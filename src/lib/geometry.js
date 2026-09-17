// Table geometry.
//
// PowerPoint.TableCell exposes NO left/top/width/height -- only the row's height
// and the column's width. So every cell rectangle here is derived: the table
// shape's origin plus the cumulative widths and heights before it.
//
// Two things to know about the numbers:
//   * row.height is the AUTHORED minimum; row.currentHeight is what is actually
//     on screen after PowerPoint grew the row to fit its text. We want the
//     latter, and fall back to the former if it is not loaded.
//   * the sums rarely match the shape's own width/height to the point, so we
//     scale them to fit. Without that, hit-testing a dropped picture drifts
//     towards the bottom-right of the table.

export const TABLE_TYPE = 'Table';

/** Read everything needed to compute cell rectangles. One sync. */
export async function readTable(context, shape) {
  const table = shape.getTable();
  table.load('rowCount,columnCount,values');
  const rows = table.rows;
  const cols = table.columns;
  rows.load('items/height,items/currentHeight,items/rowIndex');
  cols.load('items/width,items/columnIndex');
  shape.load('id,name,left,top,width,height');
  await context.sync();

  const rowHeights = rows.items
    .slice()
    .sort((a, b) => a.rowIndex - b.rowIndex)
    .map((r) => num(r.currentHeight) ?? num(r.height) ?? 0);

  const colWidths = cols.items
    .slice()
    .sort((a, b) => a.columnIndex - b.columnIndex)
    .map((c) => num(c.width) ?? 0);

  return {
    table,
    shape,
    rowCount: table.rowCount,
    columnCount: table.columnCount,
    values: table.values,
    frame: { left: shape.left, top: shape.top, width: shape.width, height: shape.height },
    rowHeights: fitTo(rowHeights, shape.height),
    colWidths: fitTo(colWidths, shape.width)
  };
}

/** Rectangle of one cell, in slide points. */
export function cellRect(geom, rowIndex, columnIndex) {
  const x = geom.frame.left + sumBefore(geom.colWidths, columnIndex);
  const y = geom.frame.top + sumBefore(geom.rowHeights, rowIndex);
  return {
    x,
    y,
    w: geom.colWidths[columnIndex] ?? 0,
    h: geom.rowHeights[rowIndex] ?? 0
  };
}

/**
 * Which cell is a floating shape sitting on? Used for "drag the picture over
 * the cell, then pull it in" -- PowerPoint gives add-ins no drop event, so we
 * work it out from where the user left the picture.
 *
 * Matches on the shape's centre point, then falls back to whichever cell it
 * overlaps most, so a picture larger than the cell still resolves sensibly.
 */
export function cellAt(geom, shapeRect) {
  const cx = shapeRect.left + shapeRect.width / 2;
  const cy = shapeRect.top + shapeRect.height / 2;

  let best = null;
  let bestOverlap = 0;

  for (let r = 0; r < geom.rowCount; r++) {
    for (let c = 0; c < geom.columnCount; c++) {
      const rect = cellRect(geom, r, c);
      if (cx >= rect.x && cx <= rect.x + rect.w && cy >= rect.y && cy <= rect.y + rect.h) {
        return { rowIndex: r, columnIndex: c, rect, byCentre: true };
      }
      const ov = overlap(rect, shapeRect);
      if (ov > bestOverlap) {
        bestOverlap = ov;
        best = { rowIndex: r, columnIndex: c, rect, byCentre: false };
      }
    }
  }
  return bestOverlap > 0 ? best : null;
}

function overlap(rect, s) {
  const w = Math.min(rect.x + rect.w, s.left + s.width) - Math.max(rect.x, s.left);
  const h = Math.min(rect.y + rect.h, s.top + s.height) - Math.max(rect.y, s.top);
  return w > 0 && h > 0 ? w * h : 0;
}

function sumBefore(arr, i) {
  let t = 0;
  for (let k = 0; k < i && k < arr.length; k++) t += arr[k];
  return t;
}

/** Scale a set of measurements so they add up to the shape's real extent. */
function fitTo(values, total) {
  const sum = values.reduce((a, b) => a + b, 0);
  if (!(sum > 0) || !(total > 0)) return values;
  const k = total / sum;
  // Only correct real drift; leave sane numbers alone.
  return Math.abs(k - 1) > 0.002 ? values.map((v) => v * k) : values;
}

function num(v) {
  return typeof v === 'number' && isFinite(v) ? v : null;
}
