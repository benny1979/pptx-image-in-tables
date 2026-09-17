// Canvas compositing.
//
// This is the trick the whole add-in rests on.
//
// PowerPoint's cell picture fill always STRETCHES the image to fill the cell,
// which normally wrecks the aspect ratio. So instead of handing PowerPoint the
// image, we hand it a transparent canvas whose aspect ratio already matches the
// cell exactly, with the image drawn into the right part of it. Stretching that
// to the cell is then an identity transform: the image lands at its true ratio,
// in the right place, and -- because it is genuinely the cell's fill -- it moves
// and resizes with the cell on any version of PowerPoint, add-in or not.

const MAX_CANVAS_PIXELS = 4e6;   // keeps the base64 payload sane
const MAX_CANVAS_EDGE = 6000;

/** Read a File into an HTMLImageElement. */
export function loadImageFromFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error('Could not read that file.'));
    reader.onload = () => resolve(reader.result);
    reader.readAsDataURL(file);
  }).then(loadImageFromDataUrl);
}

export function loadImageFromDataUrl(dataUrl) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error('That file is not an image PowerPoint can use.'));
    img.src = dataUrl;
  });
}

export function loadImageFromBase64(base64, mime = 'image/png') {
  return loadImageFromDataUrl(`data:${mime};base64,${stripDataUrl(base64)}`);
}

/**
 * Draw `img` into a transparent canvas shaped like the cell.
 *
 * @param {HTMLImageElement} img
 * @param {{w:number,h:number}} cell   cell size in points
 * @param {{x,y,w,h}} rect             image rect in points, relative to the cell
 * @returns {string} bare base64 PNG (no data: prefix -- that is what setImage wants)
 */
export function compositeToCell(img, cell, rect) {
  const scale = pickScale(img, cell, rect);
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(cell.w * scale));
  canvas.height = Math.max(1, Math.round(cell.h * scale));

  const ctx = canvas.getContext('2d');
  ctx.clearRect(0, 0, canvas.width, canvas.height);   // transparent everywhere else
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = 'high';
  ctx.drawImage(
    img,
    rect.x * scale, rect.y * scale,
    rect.w * scale, rect.h * scale
  );

  return stripDataUrl(canvas.toDataURL('image/png'));
}

/**
 * Pixels per point. We want the image drawn at no worse than its own
 * resolution, but the canvas covers the WHOLE cell, so a small image in a big
 * cell can blow the pixel budget. Hence the caps.
 */
function pickScale(img, cell, rect) {
  const wanted = rect.w > 0 ? img.naturalWidth / rect.w : 2;
  let scale = Math.min(Math.max(wanted, 1.5), 8);

  const edge = Math.max(cell.w, cell.h) * scale;
  if (edge > MAX_CANVAS_EDGE) scale *= MAX_CANVAS_EDGE / edge;

  const px = cell.w * scale * cell.h * scale;
  if (px > MAX_CANVAS_PIXELS) scale *= Math.sqrt(MAX_CANVAS_PIXELS / px);

  return Math.max(0.5, scale);
}

export function stripDataUrl(s) {
  return typeof s === 'string' ? s.replace(/^data:[^;]+;base64,/, '') : s;
}

export function approxBase64Bytes(b64) {
  return Math.floor(stripDataUrl(b64).length * 0.75);
}
