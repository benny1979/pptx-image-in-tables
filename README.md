# Cell Images

A PowerPoint add-in that puts images **inside table cells**, with the cell's text
reserved to the left, right, above or below — the thing PowerPoint tables cannot
do and Word tables can.

## How it works, and why it works that way

PowerPoint has no text-wrap engine, and `TableCell` exposes no geometry. So this
add-in does not try to reflow anything. Two mechanisms do all the work, and both
are plain OOXML that every version of PowerPoint since 2007 renders natively:

**1. The image becomes the cell's picture fill.** A cell fill always stretches to
fill the cell, which would normally wreck the aspect ratio. So the add-in
composites the image onto a transparent canvas whose aspect ratio already matches
the cell exactly, with the image drawn into the right part of it. Stretching
*that* to the cell is an identity transform: correct ratio, correct position.

Because the image genuinely is the cell's fill, it moves and resizes with the
cell, on any machine, with or without this add-in installed.

**2. Cell margins reserve the space.** Setting `margins.left` to the width of the
image band keeps PowerPoint's own text layout off the image. We never lay out
text ourselves, so nothing breaks when fonts substitute.

The consequence worth understanding: a reserved band runs the full height (or
width) of the cell. Text does **not** close in underneath a short image the way
Word's tight wrap does. That is the one visible difference from Word, and it is
a deliberate trade for never producing a layout that falls apart elsewhere.

## Requirements

| Feature | Needs | Which means |
|---|---|---|
| Everything core | PowerPointApi **1.9** | Microsoft 365: Windows build 2508+, Mac 16.100+, or PowerPoint on the web |
| Pulling in an existing picture | PowerPointApi **1.10** | Windows build 2601+, Mac 16.105+ — without it the pane asks you for the file instead |

Not available on volume-licensed perpetual or LTSC Office, and not on iPad. The
add-in checks at runtime and tells you what is missing rather than failing to
load.

**Decks you produce open correctly anywhere.** The add-in is an authoring tool;
its output is an ordinary PowerPoint table. Old perpetual Office opens, presents
and prints it fine.

## Setting it up

```bash
npm install
npm run certs      # installs a local CA so Office will trust https://localhost:3000
npm start          # leave this running
```

Then sideload `manifest.xml`:

- **Windows desktop** — put `manifest.xml` in a folder, share that folder, then
  add the share path under File ▸ Options ▸ Trust Center ▸ Trust Center Settings
  ▸ Trusted Add-in Catalogs. Restart PowerPoint; the add-in appears under
  Insert ▸ My Add-ins ▸ Shared Folder.
- **Mac** — copy `manifest.xml` into
  `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/`.
- **PowerPoint on the web** — Insert ▸ Add-ins ▸ Upload My Add-in.

The add-in then sits on the Home tab as **Cell Images**.

## Using it

1. Click a table on the slide. The pane draws its grid — pick the target cell
   there. (Office.js has no selected-cell API, and `TextRange.getParentTextFrame()`
   resolves to the table shape, not the cell, so the pane cannot read your cursor
   position. The grid is the workaround.)
2. Choose an arrangement (see below).
3. **Insert an image…** to pick a file, or drag a picture that is already on the
   slide over the cell and press **Pull in the selected picture**. PowerPoint
   gives add-ins no drop target and no shape-moved event, so that second step is
   a button rather than a real drop.
4. **Grow the cell to fit the image** widens the column (or heightens the row)
   instead of shrinking the image.

Clicking a cell that already holds an image loads its settings back into the
pane, and changing anything then moves the image straight away — no need to
re-pick the file. That needs the stored original, so it does not work on cells
placed with *Remember the original* switched off.

### Arrangements

Four bands, each with three cross-axis alignments — twelve arrangements:

| Placement | Text goes | Alignment moves the image |
|---|---|---|
| Image left | right of it, full cell height | top / centre / bottom |
| Image right | left of it, full cell height | top / centre / bottom |
| Image above | below it, full cell width | left / centre / right |
| Image below | above it, full cell width | left / centre / right |

So "image in the top-left corner, text to its right" is *image left* with
alignment *start*. Corners are not separate placements because they are already
covered.

Two more that reserve nothing, for cells where text is not competing with the
image:

| Placement | What it does |
|---|---|
| Image only | Fits the whole image inside the cell, centred. Any text overlaps it. |
| Text over image | Fills the cell edge to edge, cropping the overflow, with the text on top. |

Text alignment within whatever room is left is set separately — top / middle /
bottom and left / centre / right / justified.

**What is not possible:** text on two sides of one image. A cell has a single
text body and margins are rectangular, so reserving two bands leaves the text in
the remaining corner rectangle rather than flowing round the image. It looks
broken, so it is not offered.

### Re-fit

Editing a deck on a machine without the add-in can make a row grow, which
stretches the fill with it. **Re-fit table** recomputes every image the add-in
placed against the table's current geometry and puts it right.

This needs the original image, which is kept in a presentation-level custom XML
part (the composited fill cannot be read back — `ShapeFill` has no getter). That
roughly doubles the storage each image costs in the `.pptx`; untick *Remember the
original* if a deck gets fat, at the cost of Re-fit for those cells.

## Known limits

- No tight wrap, and no text on two sides of an image. See above — deliberate
  omissions, not gaps.
- Row height is a **minimum** in PowerPoint, so "shrink the cell to the image"
  only ever works horizontally. Vertical requests below the text's needs are
  silently ignored by PowerPoint.
- One image per cell.
- Merged cells: the add-in targets the anchor cell and reports an error for the
  covered ones.
- Undo does not compose. PowerPoint's undo stack knows nothing about the add-in's
  bookkeeping, so Ctrl+Z after an insert can leave the stored record behind. Use
  **Clear cell**.
- The pane refreshes on selection change, which is the only signal Office.js
  offers. There is no event for a table being edited, so **Refresh** exists.

## Layout

```
manifest.xml          add-in manifest (no 1.9 gate — runtime check instead)
server.cjs            HTTPS static server for sideloading
src/lib/layout.js     placement maths, six arrangements — pure, unit-tested
src/lib/compose.js    the aspect-ratio-preserving canvas composite
src/lib/geometry.js   derives cell rectangles; hit-tests a dropped picture
src/lib/store.js      custom XML part persistence
src/lib/ppt.js        everything that touches Office.js
src/taskpane/         the pane
test/                 node --test; covers layout and geometry
```

`npm test` runs the suite. It covers the maths, which is the part that can be
tested without a PowerPoint to run against — the Office.js calls cannot be, and
have not been.
