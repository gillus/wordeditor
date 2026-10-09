// Selecting across paragraphs: deleting, typing over, splitting, cutting, formatting and styling a
// multi-paragraph selection, the edits we refuse, and undo of all of it.
import { execFileSync } from 'node:child_process';
import path from 'node:path';
import assert from 'node:assert/strict';
import { withApp, caret as caretAt, select as selectIn, paste, download, view as viewOf, root, sample, outDir } from './helpers.mjs';

await withApp(async (page, errors) => {
  const caret = (...args) => caretAt(page, ...args);
  const select = (...args) => selectIn(page, ...args);
  const view = () => viewOf(page);
  const undo = () => page.keyboard.press('Control+z');
  const redo = () => page.keyboard.press('Control+Shift+z');
  const texts = async () => (await view()).map((p) => p.text);
  const para = async (prefix) => (await view()).find((p) => p.text.startsWith(prefix));
  const initial = await view();
  const assertInitial = async (msg) => assert.deepEqual(await view(), initial, msg);
  const status = () => page.locator('#status').textContent();

  // Delete across paragraphs; the hyperlink in the last one survives.
  await select(['This is ', 5], ['See ', 4]);
  await page.keyboard.press('Delete');
  assert.ok((await texts()).includes('This the website for details.'));
  assert.equal((await view()).length, initial.length - 1);
  assert.equal(await page.locator('.atom-hyperlink').textContent(), 'the website');
  await page.keyboard.type('X');
  assert.ok((await texts()).includes('This Xthe website for details.'), 'caret at the join');
  await undo();
  await undo();
  await assertInitial('undo cross delete');
  await redo();
  assert.ok((await texts()).includes('This the website for details.'), 'redo cross delete');
  await undo();

  // Typing or Enter replaces a multi-paragraph selection, as one undo step.
  await select(['First bullet', 6], ['Second bullet', 6]);
  await page.keyboard.type('Z');
  assert.ok((await texts()).includes('First Z bullet'));
  await undo();
  await assertInitial('undo typing over selection');

  await select(['Step one', 4], ['Step two', 4]);
  await page.keyboard.press('Enter');
  const split = await view();
  assert.deepEqual(['Step', ' two'].map((t) => split.find((p) => p.text === t)?.marker), ['1.', '2.']);
  await undo();
  await assertInitial('undo enter over selection');

  // Deleting across a table removes the table; undo brings back a working table.
  await select(['Step two', 4], ['Page ', 2]);
  await page.keyboard.press('Backspace');
  assert.equal(await page.locator('.tbl').count(), 0);
  assert.ok(await para('Stepge 1 of the report.'));
  await undo();
  await assertInitial('undo delete across table');
  await redo();
  assert.equal(await page.locator('.tbl').count(), 0, 'redo delete across table');
  await undo();
  await caret('R1C1');
  await page.keyboard.type('Q');
  assert.ok((await texts()).includes('R1C1Q'), 'restored table is editable');
  await undo();
  await assertInitial('undo typing in restored table');

  // Refused: selections including read-only content. Nothing changes and the user is told why.
  await select(['Page ', 0], ['Before the break.', 3]);
  await page.keyboard.press('Delete');
  await page.keyboard.type('Q');
  await assertInitial('read-only paragraph blocks the deletion');
  assert.match(await status(), /read-only/);
  await caret('Last paragraph.');
  await page.keyboard.press('Control+a');
  await page.keyboard.press('Delete');
  await assertInitial('select-all with read-only content is refused');

  // A selection from the very start of a paragraph removes it whole: what follows keeps its formatting.
  await select(['Quarterly Report', 0], ['This is ', 0]);
  await page.keyboard.press('Delete');
  assert.deepEqual((await view())[0], initial[1], 'heading removed, next paragraph unchanged');
  await undo();
  await select(['Step one', 0], ['Step two', 0]);
  await page.keyboard.press('Delete');
  assert.equal((await para('Step two')).marker, '1.', 'list item keeps its numbering');
  await undo();
  await assertInitial('undo whole-paragraph deletes');

  // Backspace into an empty paragraph removes it rather than restyling the current one.
  await caret('Quarterly Report');
  await page.keyboard.press('Enter');
  await caret('This is ', 0);
  await page.keyboard.press('Backspace');
  await assertInitial('empty paragraph removed');
  await undo();
  await undo();
  await assertInitial('undo enter + backspace');

  // Both paragraphs of the merged table cell.
  await select(['R0C0', 0], ['R0C1', 4]);
  await page.keyboard.press('Delete');
  assert.equal(await page.locator('.tbl td').first().textContent(), '');
  assert.equal(await page.locator('.tbl td').first().locator('.para').count(), 1);
  await undo();
  await assertInitial('undo delete within a cell');

  // Cut puts one line per paragraph on the clipboard; pasting it back restores the text.
  await page.evaluate(() => window.addEventListener('cut', (e) => { window.cutText = e.clipboardData.getData('text/plain'); }));
  await select(['Bookmarked', 4], ['First bullet', 6]);
  await page.keyboard.press('Control+x');
  assert.equal(await page.evaluate(() => window.cutText), 'marked paragraph.\nFirst ');
  assert.ok((await texts()).includes('Bookbullet'));
  await paste(page, await page.evaluate(() => window.cutText));
  assert.ok((await texts()).includes('Bookmarked paragraph.') && (await texts()).includes('First bullet'));
  await undo();
  await undo();
  await assertInitial('undo cut and paste');

  // Formatting and paragraph styles apply across paragraphs, as one step each.
  await select(['First bullet', 6], ['Second bullet', 6]);
  await page.keyboard.press('Control+b');
  assert.equal((await para('First bullet')).runs, 'First |400|normal/bullet|700|normal');
  assert.equal((await para('Second bullet')).runs, 'Second|700|normal/ bullet|400|normal');
  await page.click('.fmt[data-cmd="italic"]');
  assert.equal((await para('Second bullet')).runs, 'Second|700|italic/ bullet|400|normal');
  await undo();
  await undo();
  await assertInitial('undo formatting across paragraphs');
  await select(['Step one', 2], ['Step two', 2]);
  await page.selectOption('#style', 'Heading2');
  assert.deepEqual([(await para('Step one')).style, (await para('Step two')).style], ['Heading2', 'Heading2']);
  await undo();
  await assertInitial('undo style across paragraphs');

  // A mixed sequence: save it and check the file; then undo all and compare with the original.
  await select(['Step two', 4], ['Page ', 2]);
  await page.keyboard.press('Backspace');
  await select(['This is ', 5], ['See ', 4]);
  await page.keyboard.press('Delete');
  await select(['Quarterly Report', 0], ['This ', 0]);
  await page.keyboard.press('Delete');
  await select(['First bullet', 6], ['Second bullet', 6]);
  await page.keyboard.press('Control+b');
  await caret('Last paragraph.');
  await paste(page, 'Alpha\nBeta');
  const edited = await download(page, 'cross.docx');
  while (await page.locator('#undo').isEnabled()) await undo();
  await assertInitial('undo everything');
  const undone = await download(page, 'cross-undone.docx');

  assert.deepEqual(errors, []);
  execFileSync('python3', [path.join(root, 'tests/verify_cross.py'), sample, edited], { stdio: 'inherit' });
  execFileSync('python3', [path.join(root, 'tests/verify_same.py'), sample, undone], { stdio: 'inherit' });
  console.log(`select e2e OK (outputs in ${outDir})`);
});
