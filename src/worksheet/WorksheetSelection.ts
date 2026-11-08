type CellPosition = { readonly column: number; readonly row: number };
type CellRange = {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
};

/**
 * An immutable selection of a `Worksheet`: the active cell and the selected range containing it.
 *
 * Get it from `worksheet.selection`, set `worksheet.selection = new WorksheetSelection(...)` to select. When set, the range is
 * expanded to cover merged cells partially in it, and the active cell is moved to the top-left cell of its merged cell.
 */
class WorksheetSelection {
  /**
   * @param range Defaults to the active cell. Must contain the active cell.
   */
  constructor(activeCell: CellPosition, range?: CellRange) {
    const { column, row } = activeCell;
    const { left, top, right, bottom } = range ?? { left: column, top: row, right: column, bottom: row };

    if (column < left || column > right || row < top || row > bottom) {
      throw new RangeError('The range must contain the active cell.');
    }

    this.#activeCell = Object.freeze({ column, row });
    this.#range = Object.freeze({ left, top, right, bottom });
  }

  #activeCell: Readonly<CellPosition>;

  #range: Readonly<CellRange>;

  get activeCell(): Readonly<CellPosition> {
    return this.#activeCell;
  }

  get range(): Readonly<CellRange> {
    return this.#range;
  }

  /**
   * The corner of the range opposite to the active cell, i.e. where the selection extends to.
   */
  get focus(): CellPosition {
    const { column, row } = this.#activeCell;
    const { left, top, right, bottom } = this.#range;

    return { column: column === right ? left : right, row: row === bottom ? top : bottom };
  }

  toJSON() {
    return {
      activeCell: this.activeCell,
      range: this.range
    };
  }
}

export { type CellPosition, type CellRange, WorksheetSelection };
