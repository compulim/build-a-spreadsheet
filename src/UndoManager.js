/** @typedef {import('./workbook/Worksheet.js').CellValue} CellValue */
/** @typedef {import('./workbook/Worksheet.js').Worksheet} Worksheet */
/** @typedef {{ worksheet: Worksheet; column: number; row: number; value: CellValue; previousValue: CellValue }} CellChange */

/**
 * Records cell changes of tracked worksheets (or workbooks) for undo/redo. Changes made synchronously in the same task (e.g. a paste
 * or clearing a selection) are grouped into a single undo step.
 *
 * Dispatches `change` whenever `canUndo` or `canRedo` may have changed.
 */
export class UndoManager extends EventTarget {
  /** @type {CellChange[][]} */
  #undoStack = [];

  /** @type {CellChange[][]} */
  #redoStack = [];

  /** @type {CellChange[] | undefined} */
  #pending;

  #applying = false;

  /** @param {Event} event */
  #handleCellChange = event => {
    if (this.#applying) {
      return;
    }

    const { worksheet, column, row, value, previousValue } = /** @type {CellChange & Event} */ (event);

    if (!this.#pending) {
      const pending = (this.#pending = []);

      this.#undoStack.push(pending);
      this.#redoStack = [];
      queueMicrotask(() => {
        this.#pending === pending && (this.#pending = undefined);
      });
      this.dispatchEvent(new Event('change'));
    }

    this.#pending.push({ worksheet, column, row, value, previousValue });
  };

  get canUndo() {
    return this.#undoStack.length > 0;
  }

  get canRedo() {
    return this.#redoStack.length > 0;
  }

  /** The changes which `undo()` would revert. @returns {readonly CellChange[] | undefined} */
  get nextUndo() {
    return this.#undoStack.at(-1);
  }

  /** The changes which `redo()` would re-apply. @returns {readonly CellChange[] | undefined} */
  get nextRedo() {
    return this.#redoStack.at(-1);
  }

  /** @param {EventTarget} target A worksheet or workbook which dispatches `cellchange`. */
  track(target) {
    target.addEventListener('cellchange', this.#handleCellChange);
  }

  /** @param {EventTarget} target */
  untrack(target) {
    target.removeEventListener('cellchange', this.#handleCellChange);
  }

  /** Forgets all history. */
  reset() {
    this.#undoStack = [];
    this.#redoStack = [];
    this.#pending = undefined;
    this.dispatchEvent(new Event('change'));
  }

  /**
   * Reverts the last group of changes.
   *
   * @returns {CellChange[] | undefined} The changes that were reverted.
   */
  undo() {
    return this.#apply(this.#undoStack, this.#redoStack, true);
  }

  /**
   * Re-applies the last undone group of changes.
   *
   * @returns {CellChange[] | undefined} The changes that were re-applied.
   */
  redo() {
    return this.#apply(this.#redoStack, this.#undoStack, false);
  }

  /**
   * @param {CellChange[][]} from
   * @param {CellChange[][]} to
   * @param {boolean} isUndo
   */
  #apply(from, to, isUndo) {
    const changes = from.pop();

    if (!changes) {
      return;
    }

    this.#pending = undefined;
    this.#applying = true;

    try {
      const ordered = isUndo ? [...changes].reverse() : changes;

      for (const { worksheet, column, row, value, previousValue } of ordered) {
        worksheet.setValue(column, row, isUndo ? previousValue : value);
      }
    } finally {
      this.#applying = false;
    }

    to.push(changes);
    this.dispatchEvent(new Event('change'));

    return changes;
  }
}
