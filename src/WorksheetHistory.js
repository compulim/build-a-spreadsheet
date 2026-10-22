/** @typedef {import('./workbook/Worksheet.js').Worksheet} Worksheet */
/** @typedef {import('./workbook/Worksheet.js').CellChange} CellChange */

/**
 * Undo/redo history of a `Worksheet`, built by subscribing to its change events. Changes made synchronously in the same
 * task (e.g. a paste) are grouped into a single step. Loading from CSV clears the history.
 *
 * Steps are replayed through the public API of the worksheet.
 */
export class WorksheetHistory {
  /** @type {Worksheet} */
  #worksheet;

  /** @type {CellChange[][]} */
  #undoStack = [];

  /** @type {CellChange[][]} */
  #redoStack = [];

  /** @type {CellChange[] | undefined} */
  #pending;

  #applying = false;

  /** @param {Worksheet} worksheet */
  constructor(worksheet) {
    this.#worksheet = worksheet;

    const record = (/** @type {any} */ { x, y, value, previousValue, formatting, previousFormatting, edge, border, previousBorder, merge, previousMerge, type }) =>
      this.#record(
        type === 'cellchange'
          ? { x, y, value, previousValue }
          : type === 'formattingchange'
            ? { x, y, formatting, previousFormatting }
            : type === 'borderchange'
              ? { x, y, edge, border, previousBorder }
              : { x, y, merge, previousMerge }
      );

    for (const type of ['cellchange', 'formattingchange', 'borderchange', 'mergechange']) {
      worksheet.addEventListener(type, record);
    }

    worksheet.addEventListener('load', () => this.clear());
  }

  /** @param {CellChange} change */
  #record(change) {
    if (this.#applying) {
      return;
    }

    if (!this.#pending) {
      const pending = (this.#pending = []);

      this.#undoStack.push(pending);
      this.#redoStack = [];
      queueMicrotask(() => {
        this.#pending === pending && (this.#pending = undefined);
      });
    }

    this.#pending.push(change);
  }

  /** Whether `undo()` has anything to revert. */
  get canUndo() {
    return this.#undoStack.length > 0;
  }

  /** Whether `redo()` has anything to re-apply. */
  get canRedo() {
    return this.#redoStack.length > 0;
  }

  /**
   * Changes `undo()` (or `redo()`) would apply next, `undefined` if none.
   *
   * @param {'undo' | 'redo'} action
   * @returns {readonly CellChange[] | undefined}
   */
  peek(action) {
    return (action === 'undo' ? this.#undoStack : this.#redoStack).at(-1);
  }

  /** Clears the undo and redo stacks. */
  clear() {
    this.#undoStack = [];
    this.#redoStack = [];
    this.#pending = undefined;
  }

  /**
   * Reverts the last group of changes.
   *
   * @returns {boolean} `false` if there is nothing to undo.
   */
  undo() {
    return this.#apply(this.#undoStack, this.#redoStack, true);
  }

  /**
   * Re-applies the last undone group of changes.
   *
   * @returns {boolean} `false` if there is nothing to redo.
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
      return false;
    }

    const worksheet = this.#worksheet;

    this.#pending = undefined;
    to.push(changes);
    this.#applying = true;

    try {
      for (const change of isUndo ? [...changes].reverse() : changes) {
        const { x, y } = change;

        if ('edge' in change) {
          worksheet.setBorder({ x, y }, new Set([change.edge]), isUndo ? change.previousBorder : change.border);
        } else if ('merge' in change) {
          const merge = isUndo ? change.previousMerge : change.merge;

          merge ? worksheet.mergeCells(x, y, merge.width, merge.height) : worksheet.unmergeCells(x, y);
        } else if ('formatting' in change) {
          worksheet.setFormattings([[isUndo ? change.previousFormatting : change.formatting]], x, y);
        } else {
          worksheet.setValue(x, y, isUndo ? change.previousValue : change.value);
        }
      }
    } finally {
      this.#applying = false;
    }

    return true;
  }
}
