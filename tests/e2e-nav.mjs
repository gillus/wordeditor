// Arrow keys into, around and out of the sample's table (3x3, first two cells of row 0 merged into
// one cell holding paragraphs R0C0 and R0C1), and past the read-only paragraph.
import assert from 'node:assert/strict';
import { withApp, caret as caretAt, view as viewOf } from './helpers.mjs';

await withApp(async (page, errors) => {
  const caret = (...args) => caretAt(page, ...args);
  const press = (key) => page.keyboard.press(key);
  // Paragraph text (up to `len` chars), caret offset in it, and whether a table cell has focus.
  const where = () => page.evaluate(() => {
    const r = getSelection().getRangeAt(0);
    const para = (r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement).closest('.para');
    if (!para) return 'outside a paragraph';
    const pre = document.createRange();
    pre.setStart(para, 0);
    pre.setEnd(r.startContainer, r.startOffset);
    return `${para.textContent.slice(0, 8)}@${pre.toString().length}${document.activeElement.nodeName === 'TD' ? ' (cell)' : ''}`;
  });
  const move = async (start, offset, key, expected) => {
    await caret(start, offset);
    await press(key);
    if (expected instanceof RegExp) assert.match(await where(), expected, `${start}@${offset} + ${key}`);
    else assert.equal(await where(), expected, `${start}@${offset} + ${key}`);
  };

  // Left/Right: document order, cells row by row.
  await move('Step two', 8, 'ArrowRight', 'R0C0@0 (cell)');
  await move('R0C0', 0, 'ArrowLeft', 'Step two@8');
  await move('R0C1', 4, 'ArrowRight', 'R0C2@0 (cell)');
  await move('R1C0', 4, 'ArrowRight', 'R1C1@0 (cell)');
  await move('R1C1', 0, 'ArrowLeft', 'R1C0@4 (cell)');
  await move('R0C2', 4, 'ArrowRight', 'R1C0@0 (cell)');
  await move('R2C2', 4, 'ArrowRight', 'Page 1 o@0');
  await move('Page ', 0, 'ArrowLeft', 'R2C2@4 (cell)');
  await move('R1C0', 2, 'Control+ArrowRight', 'R1C0@4 (cell)'); // not at the edge: native word move
  await press('Control+ArrowRight');
  assert.equal(await where(), 'R1C1@0 (cell)', 'Ctrl+Right at a cell end');

  // Up/Down: the layout, at the same column.
  await move('Step two', 0, 'ArrowDown', /^R0C0@[0-3] \(cell\)$/); // list text is indented past the cell's left edge
  await move('R0C0', 0, 'ArrowUp', 'Step two@0');
  await move('R0C0', 0, 'ArrowDown', 'R0C1@0 (cell)'); // same cell: native
  await move('R0C1', 0, 'ArrowDown', 'R1C0@0 (cell)'); // the merged cell spans two columns
  await move('R1C1', 0, 'ArrowUp', 'R0C1@4 (cell)'); // column 1 is past the end of that short line
  await move('R0C2', 2, 'ArrowDown', 'R1C2@2 (cell)');
  await move('R2C0', 0, 'ArrowDown', /^Page 1 o@[01]$/); // cell padding: the cell text starts a little further right
  await move('Page ', 0, 'ArrowUp', 'R2C0@0 (cell)');

  // The column is kept through the table and out of it.
  await caret('R0C2', 2);
  for (const expected of ['R1C2@2 (cell)', 'R2C2@2 (cell)']) {
    await press('ArrowDown');
    assert.equal(await where(), expected);
  }
  await press('ArrowDown');
  const [, offset] = (await where()).match(/^Page .*@(\d+)$/) ?? assert.fail(`left the table: ${await where()}`);
  assert.ok(Number(offset) > 5, 'below column 2 of the table, not at the paragraph start');
  for (const expected of ['R2C2@2 (cell)', 'R1C2@2 (cell)', 'R0C2@2 (cell)']) {
    await press('ArrowUp');
    assert.equal(await where(), expected);
  }

  // The read-only paragraph is skipped both ways; the image paragraph after it takes the caret.
  const isImagePara = () => page.evaluate(() => {
    const n = getSelection().anchorNode;
    return !!(n.nodeType === 1 ? n : n.parentElement).closest('.para')?.querySelector('.atom-image');
  });
  await caret(' of the report.');
  await press('ArrowRight');
  assert.ok(await isImagePara(), 'Right skips the read-only paragraph');
  await press('ArrowLeft');
  assert.equal(await where(), 'Page 1 o@21', 'Left back to the end of the paragraph before');
  await press('ArrowDown');
  assert.ok(await isImagePara(), 'Down skips the read-only paragraph');

  // A cell whose text wraps: Up/Down move line by line inside it and leave from the first/last line.
  const before = await viewOf(page);
  await caret('R1C1', 4);
  await page.keyboard.type(' word'.repeat(12));
  const lines = await page.evaluate(() => {
    const td = [...document.querySelectorAll('.tbl td')].find((c) => c.textContent.startsWith('R1C1'));
    return Math.round(td.querySelector('.para').getBoundingClientRect().height / parseFloat(getComputedStyle(td.querySelector('.para')).lineHeight));
  });
  assert.ok(lines >= 3, `cell text wraps (${lines} lines)`);
  await caret('R1C1', 0);
  for (let i = 1; i < lines; i++) {
    await press('ArrowDown');
    assert.match(await where(), /^R1C1 wor@[1-9]\d* \(cell\)$/, `line ${i + 1} of the cell`);
  }
  await press('ArrowDown');
  assert.equal(await where(), 'R2C1@0 (cell)', 'Down from the last line of the cell');
  for (let i = 0; i < lines; i++) await press('ArrowUp');
  assert.equal(await where(), 'R1C1 wor@0 (cell)', 'Up through the cell to its first line');
  await press('ArrowUp');
  assert.match(await where(), /^R0C1@\d \(cell\)$/);
  // End leaves the caret at the wrap, shown at the end of its line: Down goes to the next line of the cell.
  await caret('R1C1', 0);
  for (let i = 2; i < lines; i++) await press('ArrowDown');
  await press('End');
  const lineEnd = await where();
  await press('ArrowDown');
  const next = await where();
  assert.ok(next.startsWith('R1C1 wor@') && next !== lineEnd, `End then Down stays in the cell: ${lineEnd} -> ${next}`);
  await press('ArrowDown');
  assert.match(await where(), /^R2C1@\d \(cell\)$/);
  await caret('R1C1', 0);
  await press('End');
  await press('ArrowUp');
  assert.match(await where(), /^R0C1@\d \(cell\)$/, 'Up from the end of the first line');
  await page.keyboard.press('Control+z');
  assert.deepEqual(await viewOf(page), before, 'undo typing in the cell');

  // Typing after a move edits the cell and undoes as usual.
  const initial = await viewOf(page);
  await caret('Step two', 8);
  await press('ArrowDown');
  await press('ArrowDown');
  await page.keyboard.type('Q');
  assert.ok((await viewOf(page)).some((p) => p.text === 'R0C1Q'), 'typed into the cell, at the column kept from "Step two"');
  await page.keyboard.press('Control+z');
  assert.deepEqual(await viewOf(page), initial, 'undo typing in the cell');

  assert.deepEqual(errors, []);
  console.log('nav e2e OK');
});
