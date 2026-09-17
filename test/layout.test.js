import test from 'node:test';
import assert from 'node:assert/strict';
import { layoutCell, growthFor, PLACEMENTS } from '../src/lib/layout.js';

const CELL = { cellW: 240, cellH: 120 };
const LANDSCAPE = { imgW: 1600, imgH: 900 };
const PORTRAIT = { imgW: 900, imgH: 1600 };

const ratio = (r) => r.w / r.h;

test('aspect ratio survives every placement', () => {
  for (const placement of PLACEMENTS) {
    for (const img of [LANDSCAPE, PORTRAIT]) {
      const { rect } = layoutCell({ ...CELL, ...img, placement });
      assert.ok(
        Math.abs(ratio(rect) - img.imgW / img.imgH) < 1e-9,
        `${placement} distorted a ${img.imgW}x${img.imgH} image`
      );
    }
  }
});

test('the image always stays inside the cell', () => {
  for (const placement of PLACEMENTS) {
    for (const img of [LANDSCAPE, PORTRAIT]) {
      for (const align of ['start', 'center', 'end']) {
        const { rect } = layoutCell({ ...CELL, ...img, placement, align, sizePct: 0.9 });
        assert.ok(rect.x >= -1e-9, `${placement}/${align} overflows left`);
        assert.ok(rect.y >= -1e-9, `${placement}/${align} overflows top`);
        assert.ok(rect.x + rect.w <= CELL.cellW + 1e-9, `${placement}/${align} overflows right`);
        assert.ok(rect.y + rect.h <= CELL.cellH + 1e-9, `${placement}/${align} overflows bottom`);
      }
    }
  }
});

test('the reserved margin clears the image on the placed side', () => {
  const side = { left: 'left', right: 'right', above: 'top', below: 'bottom' };
  for (const placement of PLACEMENTS) {
    const { rect, margins } = layoutCell({ ...CELL, ...LANDSCAPE, placement, gutter: 8 });
    const band = (placement === 'left' || placement === 'right') ? rect.w : rect.h;
    assert.ok(
      margins[side[placement]] >= band + 8,
      `${placement}: margin ${margins[side[placement]]} does not clear a ${band} band plus the gutter`
    );
  }
});

test('the reserved band leaves room for text', () => {
  for (const placement of PLACEMENTS) {
    const { margins } = layoutCell({ ...CELL, ...LANDSCAPE, placement, sizePct: 0.9 });
    assert.ok(CELL.cellW - margins.left - margins.right > 0, `${placement} left no width for text`);
    assert.ok(CELL.cellH - margins.top - margins.bottom > 0, `${placement} left no height for text`);
  }
});

test('a very tall image is clamped by the cell height, not the size setting', () => {
  const { rect } = layoutCell({ ...CELL, imgW: 100, imgH: 4000, placement: 'left', sizePct: 0.9 });
  assert.ok(rect.h <= CELL.cellH);
  assert.ok(rect.w < 20, 'a sliver image should stay a sliver');
});

test('growth asks for more column width sideways and more row height vertically', () => {
  const sideways = growthFor({ placement: 'left', ...CELL, ...LANDSCAPE });
  assert.ok(sideways.columnWidth > 0);
  assert.equal(sideways.rowHeight, null, 'a side image should never force a taller row');

  const stacked = growthFor({ placement: 'above', ...CELL, ...LANDSCAPE });
  assert.ok(stacked.rowHeight > 0);
});

test('bad input is rejected rather than producing NaN geometry', () => {
  assert.throws(() => layoutCell({ cellW: 0, cellH: 100, ...LANDSCAPE }), /usable size/);
  assert.throws(() => layoutCell({ ...CELL, imgW: 0, imgH: 10 }), /usable size/);
});
