import * as ppt from '../lib/ppt.js';
import * as store from '../lib/store.js';
import { loadImageFromFile, loadImageFromBase64, approxBase64Bytes } from '../lib/compose.js';

const $ = (id) => document.getElementById(id);

const state = {
  table: null,        // { shapeId, rowCount, columnCount, values, colWidths }
  cell: null,         // { rowIndex, columnIndex }
  filled: new Set(),  // "r,c" for cells we know already hold an image
  busy: false
};

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) {
    return fail('This add-in only works in PowerPoint.');
  }

  const caps = ppt.requirements();
  if (!caps.tables) {
    return fail(
      'This build of PowerPoint is missing the table APIs (PowerPointApi 1.9). ' +
      'You need Microsoft 365 on Windows build 2508 or later, Mac 16.100 or later, ' +
      'or PowerPoint on the web.'
    );
  }
  if (!caps.renderShape) {
    warn('Pulling in an existing picture needs PowerPointApi 1.10; this build will ask you for the file instead.');
  }

  $('app').classList.remove('hidden');
  wireUp();
  refreshTable();

  // The only selection signal add-ins get. There is no event for a table being
  // edited, so the pane can go stale -- hence the Refresh button as well.
  try {
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      () => refreshTable({ quiet: true })
    );
  } catch { /* older hosts: the Refresh button covers it */ }
});

function wireUp() {
  $('refresh').onclick = () => refreshTable();

  $('placement').onclick = (e) => {
    const b = e.target.closest('button[data-v]');
    if (!b) return;
    [...$('placement').children].forEach((x) => x.classList.toggle('on', x === b));
  };

  $('size').oninput = () => ($('sizeOut').textContent = `${$('size').value}%`);
  $('gutter').oninput = () => ($('gutterOut').textContent = `${$('gutter').value} pt`);

  $('insert').onclick = () => $('file').click();
  $('file').onchange = onFileChosen;
  $('adopt').onclick = () => guard(onAdopt);
  $('clear').onclick = () => guard(onClear);
  $('refit').onclick = () => guard(onRefit);
}

function options() {
  const placement = $('placement').querySelector('.on').dataset.v;
  return {
    placement,
    sizePct: Number($('size').value) / 100,
    gutter: Number($('gutter').value),
    align: $('align').value,
    textValign: $('valign').value,
    grow: $('grow').checked,
    remember: $('remember').checked
  };
}

async function refreshTable({ quiet = false } = {}) {
  try {
    const found = await PowerPoint.run(async (context) => {
      const shape = await ppt.findTable(context);
      if (!shape) return null;
      return ppt.describeTable(context, shape);
    });

    if (!found) {
      state.table = null;
      state.cell = null;
      $('tableName').textContent = 'No table selected.';
      $('grid').replaceChildren();
      return;
    }

    const changed = !state.table || state.table.shapeId !== found.shapeId;
    state.table = found;
    if (changed) state.cell = null;

    state.filled = await PowerPoint.run(async (context) => {
      const data = await store.readStore(context);
      const set = new Set();
      for (const k of Object.keys(data)) {
        const [id, rc] = k.split('!');
        if (id === String(found.shapeId)) set.add(rc);
      }
      return set;
    });

    $('tableName').textContent = `${found.name || 'Table'} — ${found.rowCount} × ${found.columnCount}`;
    drawGrid();
  } catch (e) {
    if (!quiet) say(e.message, true);
  }
}

function drawGrid() {
  const t = state.table;
  const grid = $('grid');
  const total = t.colWidths.reduce((a, b) => a + b, 0) || t.columnCount;
  grid.style.gridTemplateColumns = t.colWidths.map((w) => `${Math.max(6, (w / total) * 100)}fr`).join(' ');

  const cells = [];
  for (let r = 0; r < t.rowCount; r++) {
    for (let c = 0; c < t.columnCount; c++) {
      const b = document.createElement('button');
      const text = (t.values?.[r]?.[c] || '').trim();
      b.textContent = text || `R${r + 1}C${c + 1}`;
      b.title = text || `Row ${r + 1}, column ${c + 1}`;
      b.classList.toggle('on', state.cell?.rowIndex === r && state.cell?.columnIndex === c);
      b.classList.toggle('has', state.filled.has(`${r},${c}`));
      b.onclick = () => {
        state.cell = { rowIndex: r, columnIndex: c };
        drawGrid();
        say(`Cell R${r + 1}C${c + 1} selected.`);
      };
      cells.push(b);
    }
  }
  grid.replaceChildren(...cells);
}

function needCell() {
  if (!state.table) throw new Error('Click a table on the slide, then Refresh.');
  if (!state.cell) throw new Error('Pick a cell in the grid above first.');
  return { ...state.cell, shapeId: state.table.shapeId };
}

async function onFileChosen(e) {
  const file = e.target.files?.[0];
  e.target.value = '';   // so choosing the same file twice still fires
  if (!file) return;

  await guard(async () => {
    const { shapeId, rowIndex, columnIndex } = needCell();
    say('Reading the image…');
    const img = await loadImageFromFile(file);
    const base64 = await fileToBase64(file);
    say('Compositing and filling the cell…');
    const res = await ppt.applyImage(shapeId, rowIndex, columnIndex, img, base64, options());
    afterApply(rowIndex, columnIndex, res, file.name);
  });
}

async function onAdopt() {
  const hit = await ppt.adoptSelectedPicture();

  if (hit.needsFile) {
    say(
      `That picture belongs in R${hit.rowIndex + 1}C${hit.columnIndex + 1}, but this build ` +
      `cannot read a picture's pixels. Pick the original file and it will go in there.`
    );
    state.cell = { rowIndex: hit.rowIndex, columnIndex: hit.columnIndex };
    drawGrid();
    $('file').click();
    return;
  }

  say('Pulling the picture into the cell…');
  const img = await loadImageFromBase64(hit.base64);
  const res = await ppt.applyImage(
    hit.shapeId, hit.rowIndex, hit.columnIndex, img, hit.base64, options()
  );
  await ppt.deleteShape(hit.pictureId);

  state.cell = { rowIndex: hit.rowIndex, columnIndex: hit.columnIndex };
  afterApply(hit.rowIndex, hit.columnIndex, res, 'the selected picture');
  if (!hit.byCentre) {
    say(`${$('status').textContent} It was only overlapping, so I used the cell it covered most.`);
  }
}

async function onClear() {
  const { shapeId, rowIndex, columnIndex } = needCell();
  await ppt.clearCell(shapeId, rowIndex, columnIndex);
  state.filled.delete(`${rowIndex},${columnIndex}`);
  drawGrid();
  say(`Cleared R${rowIndex + 1}C${columnIndex + 1}.`);
}

async function onRefit() {
  if (!state.table) throw new Error('Click a table on the slide, then Refresh.');
  say('Re-fitting every remembered image to the table as it is now…');
  const { done, skipped } = await ppt.refit(state.table.shapeId);
  if (!done.length && !skipped.length) return say('Nothing in this table was placed by the add-in.');
  const parts = [`Re-fitted ${done.length} cell${done.length === 1 ? '' : 's'}.`];
  if (skipped.length) parts.push(`No stored original for ${skipped.join(', ')} — re-insert those by hand.`);
  say(parts.join(' '));
}

function afterApply(rowIndex, columnIndex, res, what) {
  state.filled.add(`${rowIndex},${columnIndex}`);
  drawGrid();
  const kb = Math.round(res.bytes / 1024);
  const note = res.remembered ? '' : ' Too big to remember, so Re-fit will ask for it again.';
  say(
    `Placed ${what} in R${rowIndex + 1}C${columnIndex + 1} ` +
    `(${res.image.w.toFixed(0)} × ${res.image.h.toFixed(0)} pt, ${kb} KB fill).${note}`
  );
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onerror = () => reject(new Error('Could not read that file.'));
    r.onload = () => resolve(String(r.result).replace(/^data:[^;]+;base64,/, ''));
    r.readAsDataURL(file);
  });
}

async function guard(fn) {
  if (state.busy) return;
  state.busy = true;
  document.body.style.cursor = 'progress';
  try {
    await fn();
  } catch (e) {
    say(e?.message || String(e), true);
  } finally {
    state.busy = false;
    document.body.style.cursor = '';
  }
}

function say(msg, isError = false) {
  const el = $('status');
  el.textContent = msg;
  el.classList.toggle('error', isError);
  if (isError) console.error(msg);
}

function warn(msg) {
  const b = $('banner');
  b.textContent = msg;
  b.classList.remove('hidden', 'error');
}

function fail(msg) {
  const b = $('banner');
  b.textContent = msg;
  b.classList.remove('hidden');
  b.classList.add('error');
}
