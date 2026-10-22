/** @typedef {import('./Worksheet.js').Worksheet} Worksheet */
/** @typedef {import('./Worksheet.js').CellChange} CellChange */
/** @typedef {import('../workbook/private/WorksheetChangeEvent.js').WorksheetChangeEvent} WorksheetChangeEvent */

/**
 * Undo/redo history of a `Worksheet`, built by subscribing to its `change` event. Each `edit()` transaction is a single
 * step. Column width changes are not recorded. Loading from CSV clears the history.
 *
 * Steps are replayed in a transaction through `Worksheet#edit()`.
 */
export class WorksheetHistory {
  /** @type {Worksheet} */
  #worksheet;

  /** @type {CellChange[][]} */
  #undoStack = [];

  /** @type {CellChange[][]} */
  #redoStack = [];

  #applying = false;

  /** @param {Worksheet} worksheet */
  constructor(worksheet) {
    this.#worksheet = worksheet;

    worksheet.addEventListener('change', event => {
      const { changes, load } = /** @type {WorksheetChangeEvent} */ (event);

      load ? this.clear() : this.#record(/** @type {CellChange[]} */ (changes.filter(({ type }) => type !== 'columnwidth')));
    });
  }

  /** @param {CellChange[]} changes */
  #record(changes) {
    if (!this.#applying && changes.length) {
      this.#undoStack.push(changes);
      this.#redoStack = [];
    }
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

    this.#applying = true;

    try {
      worksheet.edit(edit => {
        for (const change of isUndo ? [...changes].reverse() : changes) {
          const { x, y } = change;

          if (change.type === 'border') {
            edit.setBorder({ x, y }, new Set([change.edge]), isUndo ? change.previousBorder : change.border);
          } else if (change.type === 'merge') {
            const merge = isUndo ? change.previousMerge : change.merge;

            merge ? edit.mergeCells(x, y, merge.width, merge.height) : edit.unmergeCells(x, y);
          } else if (change.type === 'formatting') {
            edit.setFormattings([[isUndo ? change.previousFormatting : change.formatting]], x, y);
          } else {
            edit.setValue(x, y, isUndo ? change.previousValue : change.value);
          }
        }
      });
    } catch (error) {
      from.push(changes);

      throw error;
    } finally {
      this.#applying = false;
    }

    to.push(changes);

    return true;
  }
}
