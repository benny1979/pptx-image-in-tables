// Placement maths. Pure functions, no Office.js — so this is the part that can
// actually be unit-tested off-client.
//
// Everything here is in POINTS, which is what the PowerPoint JS API uses for
// table row heights, column widths and cell margins.

// Four bands, plus two arrangements that reserve nothing.
//
// The four bands times the three cross-axis alignments give the twelve
// arrangements that matter -- "image in the top-left corner with text to its
// right" is placement 'left' with align 'start', not a placement of its own.
//
// What is NOT here, because PowerPoint cannot represent it: text on two sides of
// one image. A cell has a single text body and margins are rectangular, so
// reserving two bands leaves text in the remaining corner rectangle rather than
// wrapping round the image. It looks broken, so it is not offered.
export const PLACEMENTS = ['left', 'right', 'above', 'below', 'centre', 'behind'];

// Arrangements that do not push the text anywhere.
const RESERVES_NOTHING = new Set(['centre', 'behind']);

export const DEFAULTS = {
  placement: 'left',
  sizePct: 0.4,     // image takes this fraction of the cell's usable width/height
  gutter: 6,        // points of clear space between the image and the text
  inset: 4,         // points between the image and the cell edge
  align: 'start',   // cross-axis: start | center | end
  textValign: 'Top' // Top | Middle | Bottom (PowerPoint.TextVerticalAlignment)
};

const isHorizontal = (placement) => placement === 'left' || placement === 'right';

/**
 * Work out where the image sits inside the cell, and how much margin the cell
 * needs so PowerPoint's own text layout keeps clear of it.
 *
 * The image rect is relative to the cell's top-left corner.
 *
 * @param {object} o
 * @param {number} o.cellW   cell width in points
 * @param {number} o.cellH   cell height in points
 * @param {number} o.imgW    natural image width in pixels
 * @param {number} o.imgH    natural image height in pixels
 * @returns {{rect: {x,y,w,h}, margins: {left,right,top,bottom}}}
 */
export function layoutCell(o) {
  const { cellW, cellH, imgW, imgH } = o;
  const placement = o.placement ?? DEFAULTS.placement;
  const sizePct = clamp(o.sizePct ?? DEFAULTS.sizePct, 0.05, 0.95);
  const gutter = o.gutter ?? DEFAULTS.gutter;
  const inset = o.inset ?? DEFAULTS.inset;
  const align = o.align ?? DEFAULTS.align;

  if (!(cellW > 0 && cellH > 0)) throw new Error('Cell has no usable size.');
  if (!(imgW > 0 && imgH > 0)) throw new Error('Image has no usable size.');

  const ratio = imgW / imgH;
  const availW = Math.max(1, cellW - inset * 2);
  const availH = Math.max(1, cellH - inset * 2);

  // 'behind' covers the whole cell and lets the overflow be cropped by the
  // canvas -- the text then sits over the image. 'centre' fits the image inside
  // the cell without moving the text, for cells that hold an image and nothing
  // else.
  if (RESERVES_NOTHING.has(placement)) {
    const cover = placement === 'behind';
    const k = cover
      ? Math.max(cellW / imgW, cellH / imgH)
      : Math.min(availW / imgW, availH / imgH);
    const w0 = imgW * k;
    const h0 = imgH * k;
    return {
      rect: { x: (cellW - w0) / 2, y: (cellH - h0) / 2, w: w0, h: h0 },
      margins: { left: inset, right: inset, top: inset, bottom: inset },
      clamped: false   // these placements ignore the size slider by design
    };
  }

  let w, h, clamped = false;
  if (isHorizontal(placement)) {
    w = availW * sizePct;
    h = w / ratio;
    if (h > availH) { h = availH; w = h * ratio; clamped = true; }   // never taller than the cell
  } else {
    h = availH * sizePct;
    w = h * ratio;
    if (w > availW) { w = availW; h = w / ratio; clamped = true; }   // never wider than the cell
  }
  // `clamped` means the cell's other axis, not the size slider, decided how
  // big the image is -- so dragging the slider changes nothing visible. That
  // is indistinguishable from a broken control unless the UI says so.

  // Position within the cell.
  let x, y;
  if (isHorizontal(placement)) {
    x = placement === 'left' ? inset : cellW - inset - w;
    y = crossAxis(align, cellH, h, inset);
  } else {
    y = placement === 'above' ? inset : cellH - inset - h;
    x = crossAxis(align, cellW, w, inset);
  }

  // Reserve the band. PowerPoint's text layout honours cell margins, so this is
  // what keeps the text off the image -- we never reflow text ourselves.
  const margins = { left: inset, right: inset, top: inset, bottom: inset };
  if (placement === 'left')  margins.left = inset + w + gutter;
  if (placement === 'right') margins.right = inset + w + gutter;
  if (placement === 'above') margins.top = inset + h + gutter;
  if (placement === 'below') margins.bottom = inset + h + gutter;

  return { rect: { x, y, w, h }, margins, clamped };
}

function crossAxis(align, extent, size, inset) {
  if (align === 'center') return (extent - size) / 2;
  if (align === 'end') return extent - inset - size;
  return inset;
}

/**
 * "Grow the cell to fit the image" rather than "fit the image to the cell".
 *
 * Returns the column width / row height needed so the image can be shown at
 * `targetPt` points on its long-constrained axis while leaving the text at
 * least `minTextPt` of room.
 *
 * Note the asymmetry, which is a PowerPoint fact rather than a choice of ours:
 * row height is a MINIMUM. Asking for a shorter row than the text needs is
 * silently ignored, so shrink-to-fit only ever works horizontally.
 */
export function growthFor(o) {
  const { placement, cellW, cellH, imgW, imgH } = o;
  if (RESERVES_NOTHING.has(placement)) return { columnWidth: null, rowHeight: null };
  const gutter = o.gutter ?? DEFAULTS.gutter;
  const inset = o.inset ?? DEFAULTS.inset;
  const minTextPt = o.minTextPt ?? 48;
  const ratio = imgW / imgH;

  if (isHorizontal(placement)) {
    const imgPt = o.targetPt ?? pxToPt(imgW);
    const band = Math.min(imgPt, 600);
    return { columnWidth: inset * 2 + band + gutter + minTextPt, rowHeight: null };
  }
  const imgPtH = o.targetPt ?? pxToPt(imgH);
  const bandH = Math.min(imgPtH, 400);
  // Keep the column wide enough that the image is not clamped sideways.
  const neededW = inset * 2 + bandH * ratio;
  return {
    columnWidth: Math.max(cellW, neededW),
    rowHeight: inset * 2 + bandH + gutter + Math.min(minTextPt, cellH)
  };
}

export const pxToPt = (px) => px * 0.75;   // assumes 96 DPI source images
export const ptToPx = (pt) => pt / 0.75;

function clamp(v, lo, hi) { return Math.min(hi, Math.max(lo, v)); }
