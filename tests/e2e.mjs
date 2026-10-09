// End-to-end round trip: open the sample in a real browser, edit it, download, verify with python-docx.
import { execFileSync } from 'node:child_process';
import path from 'node:path';
import assert from 'node:assert/strict';
import { withApp, caret as caretAt, paste, download, root, sample, outDir } from './helpers.mjs';

await withApp(async (page, errors) => {
  const caret = (...args) => caretAt(page, ...args);

  // --- rendering -----------------------------------------------------------------------------
  const info = await page.evaluate(() => ({
    markers: [...document.querySelectorAll('.para[data-marker]')].map((p) => p.dataset.marker),
    locked: [...document.querySelectorAll('.para.locked')].map((p) => p.textContent),
    link: document.querySelector('.atom-hyperlink')?.textContent,
    field: document.querySelector('.atom-field')?.textContent,
    imgLoaded: document.querySelector('.atom-image img')?.naturalWidth > 0,
    cells: document.querySelectorAll('.tbl td').length,
    colspan: document.querySelector('.tbl td')?.colSpan,
    pageBreak: !!document.querySelector('.atom-break'),
  }));
  assert.deepEqual(info.markers, ['•', '•', '1.', '2.']);
  assert.deepEqual(info.locked, ['Reviewed text inserted by a reviewer']);
  assert.equal(info.link, 'the website');
  assert.equal(info.field, '1');
  assert.ok(info.imgLoaded, 'image renders');
  assert.equal(info.cells, 8);
  assert.equal(info.colspan, 2);
  assert.ok(info.pageBreak);

  // Unedited download must leave document.xml untouched.
  const pristine = await download(page, 'pristine.docx');

  // --- editing -------------------------------------------------------------------------------
  await caret(' text.');
  await page.keyboard.type(' Appended.');

  await caret(' for details.');
  await page.keyboard.type('!');

  await caret('Last paragraph.');
  await paste(page, 'Pasted A\r\nPasted B');

  await caret('bold', 0, 4);
  await page.keyboard.press('Control+b');

  await caret('Last', 0, 4);
  await page.keyboard.press('Control+b');

  await caret('Bookmarked', 10);
  await page.keyboard.press('Enter');

  await caret('Second bullet', 0);
  await page.keyboard.press('Backspace');

  await caret('Step one');
  await page.keyboard.press('Shift+Enter');
  await page.keyboard.type('continued');

  await caret('R2C2');
  await page.keyboard.press('Tab');
  await page.keyboard.type('X');

  await caret('Before the break.', 0);
  await page.selectOption('#style', 'Heading2');

  await caret('Pasted B', 0, 6);
  await page.keyboard.press('Control+i');
  await page.keyboard.press('Control+u');

  // The read-only paragraph must not accept input.
  assert.equal(await page.locator('.para.locked').getAttribute('contenteditable'), 'false');

  const edited = await download(page, 'edited.docx');
  assert.deepEqual(errors, []);

  execFileSync('python3', [path.join(root, 'tests/verify.py'), sample, pristine, edited], { stdio: 'inherit' });
  console.log(`e2e OK (outputs in ${outDir})`);
});
