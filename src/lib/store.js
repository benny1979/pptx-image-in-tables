// Persistence.
//
// We keep, per cell: the placement settings, and (optionally) the ORIGINAL
// image bytes. The original matters because the fill PowerPoint holds is the
// composited canvas, and there is no API to read a cell's fill back out. Without
// a stored original, "re-fit after the table changed shape" would have to ask
// you to find the file again.
//
// It lives in a presentation-level custom XML part rather than in shape tags:
// tag values are not sized for image payloads.
//
// Cost, stated plainly: remembering originals roughly doubles the storage each
// image takes in the .pptx. Turn it off in the pane if a deck gets fat.

const NS = 'https://localhost:3000/cell-images/v1';
const CAP_BYTES = 1.5 * 1024 * 1024;   // per image, before we give up and store params only

export const key = (shapeId, rowIndex, columnIndex) => `${shapeId}!${rowIndex},${columnIndex}`;

export async function readStore(context) {
  const part = await findPart(context);
  if (!part) return {};
  const xml = part.getXml();
  await context.sync();
  return parse(xml.value);
}

export async function writeStore(context, data) {
  const xml = serialise(data);
  const part = await findPart(context);
  if (part) {
    part.setXml(xml);
  } else {
    context.presentation.customXmlParts.add(xml);
  }
  await context.sync();
}

export async function putCell(context, k, record, originalBase64) {
  const data = await readStore(context);
  const entry = { ...record };
  if (originalBase64 && originalBase64.length * 0.75 <= CAP_BYTES) {
    entry.original = originalBase64;
  }
  data[k] = entry;
  await writeStore(context, data);
  return Boolean(entry.original);
}

export async function dropCell(context, k) {
  const data = await readStore(context);
  if (!(k in data)) return;
  delete data[k];
  await writeStore(context, data);
}

export function fits(base64) {
  return base64.length * 0.75 <= CAP_BYTES;
}

async function findPart(context) {
  const parts = context.presentation.customXmlParts.getByNamespace(NS);
  parts.load('items/id');
  await context.sync();
  return parts.items.length ? parts.items[0] : null;
}

function serialise(data) {
  // JSON inside CDATA. The payload is base64 and JSON, so the only thing that
  // can break out is the literal "]]>", which cannot occur in either.
  return `<cellImages xmlns="${NS}"><![CDATA[${JSON.stringify(data)}]]></cellImages>`;
}

function parse(xml) {
  try {
    const m = /<!\[CDATA\[([\s\S]*?)\]\]>/.exec(xml || '');
    return m ? JSON.parse(m[1]) : {};
  } catch {
    return {};   // a corrupt part should cost you re-fit, not the add-in
  }
}
