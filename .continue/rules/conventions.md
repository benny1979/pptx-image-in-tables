# Cell Images — Conventions

**This project is not like the rest of `Dev/Apps/`.** It is an Office JS
PowerPoint add-in: Node, vanilla ES modules, no Python, no NiceGUI, no venv,
no PyInstaller, no `build.bat`. The workspace-wide NiceGUI rules do not
apply. The house **writing** style still does.

## Structure

- **`src/lib/ppt.js` is the only file allowed to touch Office.js.**
  `layout.js` and `geometry.js` are pure maths and must stay that way —
  they are the only parts testable without a PowerPoint host, and that is
  the whole reason the split exists. Do not import `Office`/`PowerPoint`
  into them.
- **Points, everywhere.** Row heights, column widths and cell margins are
  points in the PowerPoint JS API; the maths is in points end to end. Don't
  introduce pixels or EMUs without converting at the boundary and saying so.
- **Manifest `<Version>` must be >= 1.0.0.0.** Office silently ignores an
  invalid manifest — no error, nothing in the ribbon, indistinguishable from
  a sideloading problem. `0.1.0.0` cost an afternoon. Run
  `npx office-addin-manifest validate manifest.xml` after any manifest edit,
  before debugging anything else.
- **Two manifests, one `<Id>`.** `manifest.xml` points at GitHub Pages and is
  what people install; `manifest.dev.xml` points at localhost and is what you
  sideload while developing. Any change to one needs the same change to the
  other — they differ only in the origin.
- No build step and no bundler. Plain ES modules served as files. Keep it
  that way unless there is a real reason.
- Tests are node's built-in runner (`node --test`), stdlib only, no
  framework. `npm test` must keep working with no `npm install`.

## Behaviour

- **Never lay out text ourselves.** The add-in reserves space with cell
  margins and lets PowerPoint do the text. That is what makes it survive
  font substitution and machines without the add-in.
- **The output must be an ordinary PowerPoint table.** The image is the
  cell's picture fill, which every PowerPoint since 2007 renders natively.
  A deck must open, present and print correctly with the add-in absent.
- **Don't offer what PowerPoint cannot represent.** Text on two sides of one
  image is omitted deliberately — a cell has one text body and rectangular
  margins, so it renders broken. Same for shrinking a row below its text:
  PowerPoint treats row height as a minimum and ignores it.
- **Runtime capability checks, not a manifest `<Requirements>` gate.** The
  add-in should load and explain what this build of PowerPoint is missing,
  rather than silently refusing to appear.

## Office.js

- **Check a method exists before building on it.** `ppt.js` called
  `context.presentation.getActiveSlideOrNullObject()`, which is not in the
  PowerPoint JS API at all. It threw on every fallback path, the throws were
  swallowed, and the pane looked like it "did nothing". Use
  `activeSlide(context)` — `getSelectedSlides().getItemAt(0)`.
- **Never swallow an error in a background handler.** The selection-change
  refresh discarded its exception, which is why a wrong method name survived
  undetected. Record it; the Diagnostics panel shows it.
- **A scripted `.click()` on a file input does not open a picker** in the
  task pane's WebView. Use a real `<label for="...">`, and keep the input
  `visually-hidden` rather than `display:none`.

- **`npm run bust` before every commit that changes `src/`.** Otherwise the
  task pane keeps running the previous build and you debug a fix that is
  already deployed. Check the Diagnostics panel's `build:` stamp before
  believing anything the pane does.

## Persistence

- Settings and originals live in a **presentation-level custom XML part**,
  not shape tags — tag values are not sized for image payloads.
- Originals are kept because **a cell's fill cannot be read back**
  (`ShapeFill` has no getter). Re-fit depends on them.
- **Changing the XML namespace orphans the stored originals in every deck
  already written.** It is `urn:cell-images:v1` — a URN, not a URL, because
  it is only ever a lookup key. If it must change again, keep the old value
  as a read-only fallback (as `LEGACY_NS` does) rather than a clean break.
