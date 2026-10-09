# Word Editor (prototype)

Open a `.docx` in the browser, make light edits, download it again. Everything runs client-side.

```sh
npm install
npm run dev      # http://localhost:5173
npm test         # builds a sample .docx, edits it in headless Chromium, verifies the saved files
npm run build    # static site in dist/ (relative paths, so it works from any URL)
```

**GitHub Pages:** `.github/workflows/pages.yml` builds and publishes `dist/` on every push to
`main`. Enable it once under the repo's Settings → Pages → Source: "GitHub Actions". The site is
public; opened documents still never leave the browser.

## How it works

A `.docx` is a zip of XML parts. The editor loads it with JSZip, renders `word/document.xml` to
HTML, and only rewrites what you touch:

- **The page is one `contenteditable`**, so selections can span paragraphs. The browser handles
  edits inside a paragraph (typing, Ctrl+B/I/U); anything that changes paragraph structure is
  handled by hand so every HTML paragraph keeps exactly one `<w:p>`: Enter, Backspace/Delete at a
  paragraph edge, and any edit of a multi-paragraph selection (delete, cut, typing over it, Enter,
  paste, IME). Read-only paragraphs are non-editable islands; tables are islands whose cells are
  editable, so a selection can't break up a table's structure.
- **Deleting across paragraphs** joins the first and last paragraph and removes everything between,
  including whole tables. A selection starting at the very beginning of a paragraph removes that
  paragraph whole, so the one after keeps its own formatting (like deleting whole paragraphs in
  Word). Bold/italic/underline and paragraph styles also apply across paragraphs. Selections that
  include read-only paragraphs, cross table cells or content controls, or cross a section break
  are refused, with the reason shown in the status bar.
- **Arrow keys move through tables like in Word.** Because each cell is its own editable island,
  the browser's own arrow keys skip cells, walk them in source order for Up/Down, or strand the
  caret between blocks. So at a paragraph's edge the editor moves the caret itself
  (`src/navigate.js`): Left/Right go to the previous/next cell in reading order (and into or out
  of the table), Up/Down go to the cell above/below from the rendered layout, so merged cells need
  no bookkeeping, keeping the column across presses. Read-only paragraphs are skipped. Moves
  within a paragraph, or between neighbouring paragraphs, stay native.
- **Unedited paragraphs are never re-serialized.** On save, edited paragraphs are rebuilt from the
  DOM; every other part of the package is copied byte-for-byte (an unedited save is identical to
  the input).
- **Runs keep their original `<w:rPr>`**, so fonts, colours, sizes etc. survive edits. Bold,
  italic and underline changes are detected from computed style and written as overrides.
- **Things we don't model become "atoms"**: hyperlinks, fields, images, footnote references, page
  breaks, content controls and equations show as non-editable inline chips and are re-inserted
  verbatim.
- **Paragraphs with tracked changes are read-only** (highlighted), since editing around revisions
  without tracking would be misleading.
- **Undo is the editor's own** (`src/history.js`), not the browser's: the native stack can't
  follow paragraphs that get split, merged or re-rendered. Each edit records the paragraphs it
  replaces (typing and style changes 1→1, split 1→2, merge 2→1); undo/redo swap them back with
  their XML and restore the caret. Typing in one paragraph is one step until a 1 s pause, a click
  or arrow-key navigation. Ctrl+Z, Ctrl+Shift+Z / Ctrl+Y and the toolbar arrows all use it.

Source map: `src/docx/` (package, styles, numbering, XML helpers), `src/render.js` (XML → HTML),
`src/serialize.js` (HTML → XML), `src/editor.js` (editing behaviour), `src/history.js` and
`src/caret.js` (undo/redo), `src/navigate.js` (arrow keys around tables), `src/main.js` (toolbar).

## Known limitations

- No pagination: the document is one continuous page. Headers, footers, footnote text, comments
  and text boxes aren't shown (but are preserved).
- Ctrl+A selects the whole document, so Ctrl+A + Delete is refused if the document has read-only
  (tracked-change) paragraphs. Formatting across a selection that includes a table is refused.
- Cut/copy across paragraphs puts plain text on the clipboard (atoms as their visible text).
- Undo history is capped at 200 steps and cleared when another file is opened.
- Undoing a merge puts back any body-level markers that sat between the two paragraphs (rare in
  Word output) after the restored second paragraph rather than between them.
- Shift+arrow selections still use the browser's movement, so extending a selection into a table
  can behave oddly (and editing across cells is refused anyway). Page Up/Down and Ctrl+Up/Down
  are native too.
- Paste is plain text only. Table rows/columns can't be added or removed.
- Bookmarks inside an edited paragraph move to its start or end.
- Layout is approximate (themes, table styles, borders and floating images are simplified).
- Only tested against python-docx output; check results in Word/LibreOffice before relying on it.
