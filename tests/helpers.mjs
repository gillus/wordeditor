// Shared setup for the browser tests: a Vite dev server, Chromium, and editing helpers.
import { createServer } from 'vite';
import { chromium } from 'playwright';
import { mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

export const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const sample = path.join(root, 'tests/fixtures/sample.docx');
export const outDir = process.env.E2E_OUT ?? mkdtempSync(path.join(tmpdir(), 'wordeditor-'));

/** Run `fn(page, errors)` against the app with the sample document open. */
export async function withApp(fn) {
  const server = await createServer({ root, logLevel: 'error', server: { port: 0 } });
  await server.listen();
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage();
    const errors = [];
    page.on('pageerror', (e) => errors.push(e.message));
    await page.goto(server.resolvedUrls.local[0]);
    await page.setInputFiles('#file', sample);
    await page.waitForSelector('.page');
    await fn(page, errors);
  } finally {
    await browser.close();
    await server.close();
  }
}

/** Focus the paragraph containing `text` and select [offset, offset + extend) of that text. */
export function caret(page, text, offset = text.length, extend = 0) {
  return page.evaluate(({ text, offset, extend }) => {
    const walker = document.createTreeWalker(document.querySelector('.page'), NodeFilter.SHOW_TEXT);
    for (let n; (n = walker.nextNode());) {
      const i = n.data.indexOf(text);
      if (i < 0) continue;
      n.parentElement.closest('[contenteditable="true"]').focus();
      getSelection().setBaseAndExtent(n, i + offset, n, i + offset + extend);
      return;
    }
    throw new Error('text not found: ' + text);
  }, { text, offset, extend });
}

/** Select from `offA` chars into the first occurrence of `textA` to `offB` chars into `textB`. */
export function select(page, [textA, offA], [textB, offB]) {
  return page.evaluate(({ textA, offA, textB, offB }) => {
    const find = (text) => {
      const walker = document.createTreeWalker(document.querySelector('.page'), NodeFilter.SHOW_TEXT);
      for (let n; (n = walker.nextNode());) {
        const i = n.data.indexOf(text);
        if (i >= 0) return [n, i];
      }
      throw new Error('text not found: ' + text);
    };
    const [a, i] = find(textA);
    const [b, j] = find(textB);
    a.parentElement.closest('[contenteditable="true"]').focus();
    getSelection().setBaseAndExtent(a, i + offA, b, j + offB);
  }, { textA, offA, textB, offB });
}

/** What the user sees: text, style, list marker and bold/italic per paragraph. */
export function view(page) {
  return page.evaluate(() => [...document.querySelectorAll('.para')].map((el) => {
    const runs = [];
    const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
    for (let n; (n = walker.nextNode());) {
      if (!n.data || n.parentElement.closest('.atom')) continue;
      const cs = getComputedStyle(n.parentElement);
      const format = `${cs.fontWeight}|${cs.fontStyle}`;
      // Adjacent text with the same formatting may be one run or several; compare it as one.
      if (runs.at(-1)?.format === format) runs.at(-1).text += n.data;
      else runs.push({ text: n.data, format });
    }
    return {
      text: el.textContent, style: el.dataset.style, marker: el.dataset.marker ?? null,
      runs: runs.map((r) => `${r.text}|${r.format}`).join('/'),
    };
  }));
}

export function paste(page, text) {
  return page.evaluate((text) => {
    const data = new DataTransfer();
    data.setData('text/plain', text);
    document.activeElement.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }));
  }, text);
}

export async function download(page, name) {
  const [file] = await Promise.all([page.waitForEvent('download'), page.click('#save')]);
  const target = path.join(outDir, name);
  await file.saveAs(target);
  return target;
}
