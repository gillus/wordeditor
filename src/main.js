import './style.css';
import { Editor } from './editor.js';

const $ = (sel) => document.querySelector(sel);
const styleSelect = $('#style');
const status = $('#status');
let fileName = 'document.docx';
let noticeUntil = 0;

const editor = new Editor($('#doc'), (state) => {
  $('#save').disabled = !state.loaded;
  for (const b of document.querySelectorAll('.fmt')) b.disabled = !state.loaded;
  $('#undo').disabled = !state.canUndo;
  $('#redo').disabled = !state.canRedo;
  if (!state.loaded) return;

  const options = state.styles.map((s) => `${s.id}\u0000${s.name}`).join('\u0001');
  if (styleSelect.dataset.options !== options) {
    styleSelect.replaceChildren(...state.styles.map((s) => new Option(s.name, s.id)));
    styleSelect.dataset.options = options;
  }
  styleSelect.disabled = state.currentStyle == null;
  styleSelect.value = state.currentStyle ?? '';

  // An edit the editor refused: explain why for a few seconds instead of the usual summary.
  if (state.notice) {
    noticeUntil = Date.now() + 4000;
    status.textContent = state.notice;
    status.classList.add('notice');
    setTimeout(() => editor.onChange(editor.state), 4000);
    return;
  }
  if (Date.now() < noticeUntil) return;
  status.classList.remove('notice');
  const locked = state.locked ? ` · ${state.locked} read-only` : '';
  status.textContent = `${fileName} · ${state.paragraphs} paragraphs${locked} · ${state.unsaved ? 'unsaved changes' : 'no unsaved changes'}`;
});

async function openFile(file) {
  if (!file) return;
  if (editor.unsaved && !confirm('Discard unsaved changes?')) return;
  try {
    await editor.open(await file.arrayBuffer());
    fileName = file.name;
    editor.onChange(editor.state);
  } catch (err) {
    console.error(err);
    alert(`Could not open ${file.name}: ${err.message}`);
  }
}

$('#file').addEventListener('change', (e) => {
  openFile(e.target.files[0]);
  e.target.value = '';
});

window.addEventListener('dragover', (e) => e.preventDefault());
window.addEventListener('drop', (e) => {
  e.preventDefault();
  openFile(e.dataTransfer.files[0]);
});

// Keep the text selection: don't let toolbar buttons take focus.
for (const button of document.querySelectorAll('.toolbar button')) {
  button.addEventListener('mousedown', (e) => e.preventDefault());
}
for (const button of document.querySelectorAll('.fmt')) {
  button.addEventListener('click', () => editor.format(button.dataset.cmd));
}
$('#undo').addEventListener('click', () => editor.undo());
$('#redo').addEventListener('click', () => editor.redo());

styleSelect.addEventListener('change', () => editor.setParagraphStyle(styleSelect.value));

$('#save').addEventListener('click', async () => {
  const blob = await editor.save();
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = fileName;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 10_000);
});

window.addEventListener('beforeunload', (e) => {
  if (editor.unsaved) e.preventDefault();
});
