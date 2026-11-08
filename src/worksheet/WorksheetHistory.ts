import type { Worksheet, WorksheetRedo, WorksheetUndo } from './Worksheet.ts';
import type { WorksheetChangeEvent } from './WorksheetChangeEvent.ts';

/**
 * Undo/redo history of a `Worksheet`. It remembers the `undo()` objects of `edit()` transactions (from the `change`
 * event) and the `redo()` objects returned by undoing them. Loading from CSV clears the history.
 */
class WorksheetHistory {
  #undoStack: WorksheetUndo[] = [];

  #redoStack: WorksheetRedo[] = [];

  constructor(worksheet: Worksheet) {
    worksheet.addEventListener('change', event => {
      const { load, undo } = event as WorksheetChangeEvent;

      if (load) {
        this.clear();
      } else if (undo) {
        this.#undoStack.push(undo);
        this.#redoStack = [];
      }
    });
  }

  /** Whether `undo()` has anything to revert. */
  get canUndo(): boolean {
    return this.#undoStack.length > 0;
  }

  /** Whether `redo()` has anything to re-apply. */
  get canRedo(): boolean {
    return this.#redoStack.length > 0;
  }

  /** Clears the undo and redo stacks. */
  clear(): void {
    this.#undoStack = [];
    this.#redoStack = [];
  }

  /**
   * Reverts the last transaction.
   *
   * @returns `false` if there is nothing to undo.
   */
  undo(): boolean {
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
   * @returns `false` if there is nothing to redo.
   */
  redo(): boolean {
    const entry = this.#redoStack.at(-1);

    if (!entry) {
      return false;
    }

    this.#undoStack.push(entry.redo());
    this.#redoStack.pop();

    return true;
  }
}

export { WorksheetHistory };
