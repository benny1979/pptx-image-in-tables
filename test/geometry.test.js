import test from 'node:test';
import assert from 'node:assert/strict';
import { cellRect, cellAt } from '../src/lib/geometry.js';

// A 3x3 table at (100, 50), 300 x 150 points, with uneven columns.
const geom = {
  rowCount: 3,
  columnCount: 3,
  frame: { left: 100, top: 50, width: 300, height: 150 },
  colWidths: [100, 150, 50],
  rowHeights: [40, 60, 50]
};

test('cell rectangles tile the table exactly', () => {
  assert.deepEqual(cellRect(geom, 0, 0), { x: 100, y: 50, w: 100, h: 40 });
  assert.deepEqual(cellRect(geom, 1, 1), { x: 200, y: 90, w: 150, h: 60 });

  const last = cellRect(geom, 2, 2);
  assert.equal(last.x + last.w, geom.frame.left + geom.frame.width);
  assert.equal(last.y + last.h, geom.frame.top + geom.frame.height);
});

test('a picture dropped over a cell resolves to that cell', () => {
  // Centre of R1C1 is (275, 120).
  const hit = cellAt(geom, { left: 265, top: 110, width: 20, height: 20 });
  assert.deepEqual([hit.rowIndex, hit.columnIndex], [1, 1]);
  assert.equal(hit.byCentre, true);
});

test('a picture bigger than the cell falls back to the biggest overlap', () => {
  // Centre (100, 40) sits above the table, but it blankets the top-left cell.
  const hit = cellAt(geom, { left: 40, top: 10, width: 120, height: 60 });
  assert.ok(hit);
  assert.equal(hit.byCentre, false);
  assert.deepEqual([hit.rowIndex, hit.columnIndex], [0, 0]);
});

test('a picture nowhere near the table resolves to nothing', () => {
  assert.equal(cellAt(geom, { left: 600, top: 600, width: 50, height: 50 }), null);
});
