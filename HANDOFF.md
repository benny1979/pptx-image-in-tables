# Cell Images — Handoff

## What it is

A **PowerPoint task-pane add-in** (Office JS) that puts an image **inside a
table cell**, with the cell's text reserved to one side — what Word tables do
and PowerPoint's cannot.

Not a NiceGUI app. Nothing in this repo resembles the rest of
`Dev/Apps/` — no Python, no venv, no `build.bat`. Node + a static HTTPS
server, sideloaded into PowerPoint.

`README.md` is the real documentation: how it works, all twelve
arrangements, requirements, sideloading, and the known limits. It is
thorough — **read it before changing anything**, and do not duplicate it
here.

## State — built and working, never click-tested in real PowerPoint

- Two commits, both **2026-09-17**. Nothing since.
- `npm test` → **14 passing** (node's built-in runner, no framework).
- The tests cover `layout.js` and `geometry.js` only — the pure maths. That
  is deliberate: they are the only parts that run without an Office host.
- **Everything that touches Office.js (`ppt.js`, `store.js`, the task pane)
  has never been exercised against a real PowerPoint.** No sideload run is
  recorded anywhere.

## Came to the Drive late

Cloned here **2026-09-25** from `benny1979/pptx-image-in-tables`, eight days
after it was written. Until then it existed **only on GitHub** — no local
copy, no `NOTES.md`, no `TASKS.md` entry. A Drive-wide search for it found
nothing, which is exactly why it went missing.

## Running it

```bash
npm install
npm run certs    # installs a local CA so Office trusts https://localhost:3000
npm start        # leave running
```

Then sideload `manifest.xml` — per-platform steps are in `README.md`.

- Node **v26.7.0** / npm **11.19.0** on this machine (2026-09-25).
- `npm install` has **not** been run in this clone; `node_modules/` is
  gitignored, so the first run here will fetch `office-addin-dev-certs`.
- `npm test` needs no install — it is stdlib only.

## Shape of the code

| File | Does |
|---|---|
| `src/lib/layout.js` | Placement maths, in **points**. Pure, tested. |
| `src/lib/geometry.js` | Table/cell rectangles, hit-testing a picture to a cell. Pure, tested. |
| `src/lib/compose.js` | Composites the image onto a transparent canvas matching the cell's aspect ratio. |
| `src/lib/ppt.js` | **Everything** that touches Office.js. The only file that needs a host. |
| `src/lib/store.js` | Per-cell settings + optional original bytes, in a presentation-level custom XML part. |
| `src/taskpane/` | The pane: the cell grid, the arrangement controls. |
| `server.cjs` | Minimal HTTPS static server for sideloading. Refuses to serve outside the project dir. |

## Things that will bite

- **`store.js`'s XML namespace is `urn:cell-images:v1`** — the label the
  hidden part is found by, never fetched. It was the dev server's address
  until 2026-09-25; `LEGACY_NS` still reads that one so any deck written
  before then keeps its stored originals. Changing the namespace again
  orphans every deck that predates the change, so don't, casually.
- **The manifest deliberately has no `<Requirements>` gate** on
  PowerPointApi 1.9 — the add-in checks at runtime and explains what is
  missing, instead of silently refusing to load. Don't "fix" this.
- **Office.js cannot tell you which cell the cursor is in.** There is no
  selected-cell API and `TextRange.getParentTextFrame()` resolves to the
  table shape. The pane's clickable grid is the workaround, not a UI whim.
- **A cell's fill cannot be read back** (`ShapeFill` has no getter), which is
  the whole reason originals are stored at all. Re-fit depends on it.
- **Storing originals roughly doubles each image's cost in the `.pptx`.**
  There is a 1.5 MB per-image cap in `store.js` after which only the
  parameters are kept.
- **Undo does not compose** — Ctrl+Z after an insert can leave the stored
  record behind. Use *Clear cell*.

## Next

1. **Sideload it and actually use it.** Everything Office-facing is
   unverified. That is the single biggest gap.
2. Drop `LEGACY_NS` once no deck written before 2026-09-25 matters —
   realistically, once the add-in has been used in anger at all.

Not related to `chalk`'s in-cell images (a canvas editor, its own model) or
to `pptx_combiner`'s table-to-picture flattening — different problems that
merely sound alike.
