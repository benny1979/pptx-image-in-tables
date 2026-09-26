// Everything that touches Office.js lives here.

import { readTable, cellRect, cellAt } from './geometry.js?v=558a912';
import { layoutCell, growthFor, DEFAULTS } from './layout.js?v=558a912';
import { compositeToCell, loadImageFromBase64, stripDataUrl } from './compose.js?v=558a912';
import * as store from './store.js?v=558a912';

/** What this build of PowerPoint can actually do. */
export function requirements() {
  const has = (v) => {
    try { return Office.context.requirements.isSetSupported('PowerPointApi', v); }
    catch { return false; }
  };
  return {
    fill: has('1.8'),        // ShapeFill.setImage
    tables: has('1.9'),      // TableCell.margins, .fill, .borders -- the core of this add-in
    renderShape: has('1.10') // Shape.getImageAsBase64, used to adopt an existing picture
  };
}

/**
 * The slide being edited.
 *
 * NOT `context.presentation.getActiveSlideOrNullObject()` -- that method does
 * not exist in the PowerPoint JS API. Calling it threw
 * "is not a function" on every fallback path, and because those throws were
 * swallowed, the pane simply appeared to do nothing. Diagnosed 2026-09-26
 * from the add-in's own diagnostics panel.
 *
 * `getSelectedSlides()` (PowerPointApi 1.5) is the real one; the slide being
 * edited is always the selected slide.
 */
export function activeSlide(context) {
  return context.presentation.getSelectedSlides().getItemAt(0);
}

const isTable = (shape) => String(shape.type || '').toLowerCase() === 'table';
const isPicture = (shape) => {
  const t = String(shape.type || '').toLowerCase();
  return t === 'image' || t === 'picture';
};

/**
 * Which table are we acting on? The selected one, or -- if nothing useful is
 * selected -- the only table on the slide.
 *
 * Note we cannot ask which CELL you are in: Office.js has no selected-cell API,
 * and TextRange.getParentTextFrame() resolves to the table shape, not the cell.
 * That is why the pane draws its own grid for you to click.
 */
export async function findTable(context) {
  const selected = context.presentation.getSelectedShapes();
  selected.load('items/id,items/type,items/name');
  await context.sync();

  const picked = selected.items.find(isTable);
  if (picked) return picked;

  const shapes = activeSlide(context).shapes;
  shapes.load('items/id,items/type,items/name');
  await context.sync();

  const tables = shapes.items.filter(isTable);
  if (tables.length === 1) return tables[0];
  if (tables.length > 1) throw new Error('More than one table on this slide -- click the one you want first.');
  return null;
}

/** Grid summary for the cell picker. */
export async function describeTable(context, shape) {
  const geom = await readTable(context, shape);
  return {
    shapeId: shape.id,
    name: shape.name,
    rowCount: geom.rowCount,
    columnCount: geom.columnCount,
    values: geom.values,
    colWidths: geom.colWidths,
    rowHeights: geom.rowHeights
  };
}

/**
 * Put an image into a cell.
 *
 * @param {HTMLImageElement} img
 * @param {object} opts  placement/sizePct/gutter/inset/align/textValign, plus
 *                       grow:boolean and remember:boolean
 */
export async function applyImage(shapeId, rowIndex, columnIndex, img, sourceBase64, opts) {
  const o = { ...DEFAULTS, ...opts };

  return PowerPoint.run(async (context) => {
    let shape = await shapeById(context, shapeId);
    let geom = await readTable(context, shape);

    // "Grow the cell to fit the image" runs first, because it changes the very
    // cell size the layout is computed against.
    if (o.grow) {
      const rect0 = cellRect(geom, rowIndex, columnIndex);
      const g = growthFor({
        placement: o.placement,
        cellW: rect0.w, cellH: rect0.h,
        imgW: img.naturalWidth, imgH: img.naturalHeight,
        gutter: o.gutter, inset: o.inset
      });
      const table = shape.getTable();
      if (g.columnWidth) {
        const col = table.columns.getItemAt(columnIndex);
        col.width = g.columnWidth;
      }
      if (g.rowHeight) {
        const row = table.rows.getItemAt(rowIndex);
        row.height = g.rowHeight;   // a MINIMUM -- PowerPoint may still grow it further
      }
      await context.sync();

      shape = await shapeById(context, shapeId);
      geom = await readTable(context, shape);
    }

    const rect = cellRect(geom, rowIndex, columnIndex);
    const { rect: imgRect, margins } = layoutCell({
      cellW: rect.w, cellH: rect.h,
      imgW: img.naturalWidth, imgH: img.naturalHeight,
      placement: o.placement, sizePct: o.sizePct,
      gutter: o.gutter, inset: o.inset, align: o.align
    });

    const composited = compositeToCell(img, { w: rect.w, h: rect.h }, imgRect);

    const table = shape.getTable();
    const cell = table.getCellOrNullObject(rowIndex, columnIndex);
    cell.load('isNullObject');
    await context.sync();
    if (cell.isNullObject) throw new Error('That cell does not exist -- it may be part of a merged area.');

    cell.fill.setImage(composited);
    cell.margins.left = margins.left;
    cell.margins.right = margins.right;
    cell.margins.top = margins.top;
    cell.margins.bottom = margins.bottom;
    cell.verticalAlignment = o.textValign;
    if (o.textHalign) cell.horizontalAlignment = o.textHalign;
    await context.sync();

    let remembered = false;
    if (o.remember !== false && sourceBase64) {
      remembered = await store.putCell(
        context,
        store.key(shapeId, rowIndex, columnIndex),
        {
          placement: o.placement, sizePct: o.sizePct, gutter: o.gutter,
          inset: o.inset, align: o.align,
          textValign: o.textValign, textHalign: o.textHalign
        },
        stripDataUrl(sourceBase64)
      );
    }

    return { cell: rect, image: imgRect, bytes: Math.floor(composited.length * 0.75), remembered };
  });
}

/** Take the image back out and put the cell's margins back to something normal. */
export async function clearCell(shapeId, rowIndex, columnIndex) {
  return PowerPoint.run(async (context) => {
    const shape = await shapeById(context, shapeId);
    const cell = shape.getTable().getCellOrNullObject(rowIndex, columnIndex);
    cell.load('isNullObject');
    await context.sync();
    if (cell.isNullObject) return;

    cell.fill.clear();
    cell.margins.left = 7.2;
    cell.margins.right = 7.2;
    cell.margins.top = 3.6;
    cell.margins.bottom = 3.6;
    await context.sync();

    await store.dropCell(context, store.key(shapeId, rowIndex, columnIndex));
  });
}

/**
 * Pull a picture that is already on the slide into whichever cell it is sitting
 * over. PowerPoint exposes no drop target and no shape-moved event to add-ins,
 * so the flow is: you drag it over the cell yourself, then press the button.
 */
export async function adoptSelectedPicture() {
  const caps = requirements();

  const found = await PowerPoint.run(async (context) => {
    const sel = context.presentation.getSelectedShapes();
    sel.load('items/id,items/type,items/name,items/left,items/top,items/width,items/height');
    await context.sync();

    const pic = sel.items.find(isPicture);
    if (!pic) throw new Error('Select the picture you want to pull in first.');

    const table = await findTableOnSlide(context);
    if (!table) throw new Error('No table on this slide to pull it into.');

    const geom = await readTable(context, table);
    const hit = cellAt(geom, { left: pic.left, top: pic.top, width: pic.width, height: pic.height });
    if (!hit) throw new Error('That picture is not over the table -- drag it onto the cell, then try again.');

    if (!caps.renderShape) {
      return { needsFile: true, shapeId: table.id, pictureId: pic.id, ...hit };
    }

    const rendered = pic.getImageAsBase64();
    await context.sync();
    return { base64: stripDataUrl(rendered.value), shapeId: table.id, pictureId: pic.id, ...hit };
  });

  return found;
}

/** Delete the floating picture once its pixels are safely in the cell. */
export async function deleteShape(shapeId) {
  return PowerPoint.run(async (context) => {
    const shape = await shapeById(context, shapeId);
    shape.delete();
    await context.sync();
  });
}

/**
 * What was this cell set to last time? Lets the pane show you the arrangement
 * you actually have, rather than resetting to defaults every time you click.
 *
 * Returns null for a cell the add-in did not place, and hasOriginal:false for
 * one placed with "remember the original" switched off -- that one can be
 * re-positioned only by supplying the file again.
 */
export async function readCellSettings(shapeId, rowIndex, columnIndex) {
  return PowerPoint.run(async (context) => {
    const data = await store.readStore(context);
    const rec = data[store.key(shapeId, rowIndex, columnIndex)];
    if (!rec) return null;
    const { original, ...settings } = rec;
    return { ...settings, hasOriginal: Boolean(original) };
  });
}

/**
 * Move the image that is already in a cell -- left to right, above to below,
 * realign it, resize it -- without asking for the file again.
 *
 * This is just applyImage against the stored original. Cheap, because the cell
 * fill is regenerated from scratch on every placement anyway.
 */
export async function reapply(shapeId, rowIndex, columnIndex, opts) {
  const original = await PowerPoint.run(async (context) => {
    const data = await store.readStore(context);
    return data[store.key(shapeId, rowIndex, columnIndex)]?.original || null;
  });

  if (!original) {
    throw new Error('No stored original for that cell, so it cannot be moved. Insert the image again.');
  }

  const img = await loadImageFromBase64(original);
  return applyImage(shapeId, rowIndex, columnIndex, img, original, { ...opts, grow: false });
}

/**
 * Recompute every remembered cell against the table's current geometry.
 *
 * This is the repair tool. Edit a deck on a machine without the add-in, let a
 * row grow, and the fill stretches with it; run this at home and it is right
 * again.
 */
export async function refit(shapeId) {
  const records = await PowerPoint.run(async (context) => {
    const data = await store.readStore(context);
    return Object.entries(data)
      .filter(([k]) => k.startsWith(`${shapeId}!`))
      .map(([k, v]) => {
        const [, rc] = k.split('!');
        const [r, c] = rc.split(',').map(Number);
        return { rowIndex: r, columnIndex: c, record: v };
      });
  });

  const done = [];
  const skipped = [];
  for (const { rowIndex, columnIndex, record } of records) {
    if (!record.original) { skipped.push(`R${rowIndex + 1}C${columnIndex + 1}`); continue; }
    const img = await loadImageFromBase64(record.original);
    await applyImage(shapeId, rowIndex, columnIndex, img, record.original, { ...record, grow: false });
    done.push(`R${rowIndex + 1}C${columnIndex + 1}`);
  }
  return { done, skipped };
}

async function shapeById(context, shapeId) {
  const shapes = activeSlide(context).shapes;
  shapes.load('items/id,items/type,items/name,items/left,items/top,items/width,items/height');
  await context.sync();
  const shape = shapes.items.find((s) => s.id === shapeId);
  if (!shape) throw new Error('Lost track of that table -- click it again.');
  return shape;
}

async function findTableOnSlide(context) {
  const shapes = activeSlide(context).shapes;
  shapes.load('items/id,items/type,items/name,items/left,items/top,items/width,items/height');
  await context.sync();
  const tables = shapes.items.filter(isTable);
  return tables.length ? tables[0] : null;
}
