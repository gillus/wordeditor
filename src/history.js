// Undo/redo stacks for paragraph-level edits.
//
// Every edit replaces a run of adjacent blocks with another run: typing or a style change is
// 1 -> 1, a split 1 -> 2, a merge 2 -> 1, deleting a selection n -> 1. Blocks are paragraphs,
// plus whole tables etc. that a deletion removed. An entry has a `before` and an `after` side:
//   els:   the HTML block elements (reused across undo/redo, so older entries stay valid)
//   ps:    detached XML nodes for that side while it is off screen (null while on screen)
//   caret: where the caret was on that side
// Undo and redo hand the two sides to `swap`, which takes one off screen and puts the other on.

const LIMIT = 200;

export class History {
  constructor(swap) {
    this.swap = swap;
    this.clear();
  }

  clear() {
    this.undoStack = [];
    this.redoStack = [];
  }

  get canUndo() {
    return this.undoStack.length > 0;
  }

  get canRedo() {
    return this.redoStack.length > 0;
  }

  get top() {
    return this.undoStack.at(-1) ?? null;
  }

  push(entry) {
    this.undoStack.push(entry);
    if (this.undoStack.length > LIMIT) this.undoStack.shift();
    this.redoStack.length = 0;
  }

  /** Stop the latest entry from absorbing further typing. */
  seal() {
    if (this.top) this.top.open = false;
  }

  undo() {
    const entry = this.undoStack.pop();
    if (!entry) return false;
    this.swap(entry.after, entry.before);
    entry.open = false;
    this.redoStack.push(entry);
    this.seal();
    return true;
  }

  redo() {
    const entry = this.redoStack.pop();
    if (!entry) return false;
    this.swap(entry.before, entry.after);
    this.undoStack.push(entry);
    return true;
  }
}
