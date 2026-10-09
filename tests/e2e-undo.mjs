// Undo/redo: every kind of edit must undo back to the exact original, and redo forward again.
import { execFileSync } from 'node:child_process';
import path from 'node:path';
import assert from 'node:assert/strict';
import { withApp, caret as caretAt, paste, download, view as viewOf, root, sample, outDir } from './helpers.mjs';

await withApp(async (page, errors) => {
  const caret = (...args) => caretAt(page, ...args);
  const view = () => viewOf(page);
  const undo = () => page.keyboard.press('Control+z');
  const redo = () => page.keyboard.press('Control+Shift+z');
  const texts = async () => (await view()).map((p) => p.text);
  const initial = await view();
  const assertInitial = async (msg) => assert.deepEqual(await view(), initial, msg);

  // Typing: a burst is one step; a pause starts a new one.
  await caret(' text.');
  await page.keyboard.type('abc');
  await undo();
  await assertInitial('undo typing burst');
  await redo();
  assert.ok((await texts()).includes('This is bold and italic text.abc'), 'redo typing');
  await undo();

  await caret(' text.');
  await page.keyboard.type('a');
  await page.waitForTimeout(1200);
  await page.keyboard.type('b');
  await undo();
  assert.ok((await texts()).includes('This is bold and italic text.a'), 'pause splits undo steps');
  await undo();
  await assertInitial('undo both typing steps');

  // The caret goes back to where the undone edit happened.
  await caret('Step two');
  await page.keyboard.type('XY');
  await undo();
  await page.keyboard.type('Z');
  assert.ok((await texts()).includes('Step twoZ'), 'caret restored after undo');
  await undo();
  await assertInitial('undo after caret check');

  // Split.
  await caret('Bookmarked', 10);
  await page.keyboard.press('Enter');
  assert.ok((await texts()).includes(' paragraph.'));
  await undo();
  await assertInitial('undo split');
  await redo();
  assert.ok((await texts()).includes('Bookmarked') && (await texts()).includes(' paragraph.'), 'redo split');
  await undo();
  await assertInitial('undo split again');

  // Merge (Backspace) across a paragraph holding a hyperlink atom, then via Delete.
  await caret('See ', 0);
  await page.keyboard.press('Backspace');
  assert.ok((await texts()).includes('This is bold and italic text.See the website for details.'));
  await undo();
  await assertInitial('undo backspace merge');
  assert.equal(await page.locator('.atom-hyperlink').textContent(), 'the website');
  await redo();
  await undo();
  await caret('First bullet');
  await page.keyboard.press('Delete');
  assert.ok((await texts()).includes('First bulletSecond bullet'));
  await undo();
  await assertInitial('undo delete merge');

  // Paragraph style.
  await caret('Before the break.', 3);
  await page.selectOption('#style', 'Heading2');
  await undo();
  await assertInitial('undo style change');
  await redo();
  assert.equal((await view()).find((p) => p.text.startsWith('Before the break.')).style, 'Heading2', 'redo style');
  await undo();

  // Formatting and multi-line paste.
  await caret('Last', 0, 4);
  await page.keyboard.press('Control+b');
  await undo();
  await assertInitial('undo bold');
  await caret('Last paragraph.');
  await paste(page, 'one\ntwo\nthree');
  assert.ok((await texts()).includes('three'));
  await undo();
  await assertInitial('one undo step for a multi-line paste');

  // Toolbar buttons, and a new edit clearing the redo stack.
  await caret('Step one');
  await page.keyboard.type('!');
  await page.click('#undo');
  await assertInitial('toolbar undo');
  assert.equal(await page.locator('#redo').isEnabled(), true);
  await page.click('#redo');
  await page.click('#undo');
  await caret('Step one');
  await page.keyboard.type('?');
  assert.equal(await page.locator('#redo').isEnabled(), false, 'new edit clears redo');
  await undo();

  // A mixed sequence: undo everything, check the saved file equals the original; redo it all.
  await caret('Bookmarked', 10);
  await page.keyboard.type(' one');
  await page.keyboard.press('Enter');
  await page.keyboard.type('two ');
  await caret('First bullet');
  await page.keyboard.press('Delete');
  await caret('Step two', 4);
  await page.selectOption('#style', 'Heading1');
  await caret('Last', 0, 4);
  await page.keyboard.press('Control+i');
  const edited = await view();
  assert.notDeepEqual(edited, initial);

  let steps = 0;
  while (await page.locator('#undo').isEnabled()) {
    await undo();
    steps++;
  }
  await assertInitial('undo everything');
  const undone = await download(page, 'undone.docx');
  while (await page.locator('#redo').isEnabled()) await redo();
  const redone = await view();
  for (let i = 0; i < Math.max(redone.length, edited.length); i++) {
    assert.deepEqual(redone[i], edited[i], `redo all ${steps} steps: paragraph ${i}`);
  }

  assert.deepEqual(errors, []);
  execFileSync('python3', [path.join(root, 'tests/verify_same.py'), sample, undone], { stdio: 'inherit' });
  console.log(`undo e2e OK (${steps} steps; outputs in ${outDir})`);
});
