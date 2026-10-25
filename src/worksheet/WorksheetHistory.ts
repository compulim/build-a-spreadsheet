/** @typedef {import('./Worksheet.ts').Worksheet} Worksheet */
/** @typedef {import('./Worksheet.ts').WorksheetUndo} WorksheetUndo */
/** @typedef {import('./Worksheet.ts').WorksheetRedo} WorksheetRedo */
/** @typedef {import('./WorksheetChangeEvent.ts').WorksheetChangeEvent} WorksheetChangeEvent */

/**
 * Undo/redo history of a `Worksheet`. It remembers the `undo()` objects of `edit()` transactions (from the `change`
 * event) and the `redo()` objects returned by undoing them. Loading from CSV clears the history.
 */
export class WorksheetHistory {
  /** @type {WorksheetUndo[]} */
  #undoStack = [];

  /** @type {WorksheetRedo[]} */
  #redoStack = [];

  /** @param {Worksheet} worksheet */
  constructor(worksheet) {
    worksheet.addEventListener('change', event => {
      const { load, undo } = /** @type {WorksheetChangeEvent} */ (event);

      if (load) {
        this.clear();
      } else if (undo) {
        this.#undoStack.push(undo);
        this.#redoStack = [];
      }
    });
  }

  /** Whether `undo()` has anything to revert. */
  get canUndo() {
    return this.#undoStack.length > 0;
  }

  /** Whether `redo()` has anything to re-apply. */
  get canRedo() {
    return this.#redoStack.length > 0;
  }

  /** Clears the undo and redo stacks. */
  clear() {
    this.#undoStack = [];
    this.#redoStack = [];
  }

  /**
   * Reverts the last transaction.
   *
   * @returns {boolean} `false` if there is nothing to undo.
   */
  undo() {
    const entry = this.#undoStack.at(-1);

    if (!entry) {
      return false;
    }

    this.#redoStack.push(entry.undo());
    this.#undoStack.pop();

    return true;
  }

  /**
   * Re-applies the last undone transaction.
   *
   * @returns {boolean} `false` if there is nothing to redo.
   */
  redo() {
    const entry = this.#redoStack.at(-1);

    if (!entry) {
      return false;
    }

    this.#undoStack.push(entry.redo());
    this.#redoStack.pop();

    return true;
  }
}
