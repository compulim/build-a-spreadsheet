/**
 * @typedef {{ column: number; row: number }} CellPosition
 * @typedef {{ left: number; top: number; right: number; bottom: number }} CellRange
 */

/**
 * An immutable selection of a `<compulim-worksheet>`: the active cell and the selected range containing it.
 *
 * Get it from `element.selection`, set `element.selection = new WorksheetSelection(...)` to select. When set, the range is
 * expanded to cover merged cells partially in it, and the active cell is moved to the top-left cell of its merged cell.
 */
export class WorksheetSelection {
  /**
   * @param {CellPosition} activeCell
   * @param {CellRange} [range] Defaults to the active cell. Must contain the active cell.
   */
  constructor(activeCell, range) {
    const { column, row } = activeCell;
    const { left, top, right, bottom } = range ?? { left: column, top: row, right: column, bottom: row };

    if (column < left || column > right || row < top || row > bottom) {
      throw new RangeError('The range must contain the active cell.');
    }

    this.#activeCell = Object.freeze({ column, row });
    this.#range = Object.freeze({ left, top, right, bottom });
  }

  /** @type {Readonly<CellPosition>} */
  #activeCell;

  /** @type {Readonly<CellRange>} */
  #range;

  /** @returns {Readonly<CellPosition>} */
  get activeCell() {
    return this.#activeCell;
  }

  /** @returns {Readonly<CellRange>} */
  get range() {
    return this.#range;
  }

  /**
   * The corner of the range opposite to the active cell, i.e. where the selection extends to.
   *
   * @returns {CellPosition}
   */
  get focus() {
    const { column, row } = this.#activeCell;
    const { left, top, right, bottom } = this.#range;

    return { column: column === right ? left : right, row: row === bottom ? top : bottom };
  }
}
