/** @typedef {import('./Worksheet.js').CellValue} CellValue */
/** @typedef {import('./Worksheet.js').Worksheet} Worksheet */

export class CellChangeEvent extends Event {
  /**
   * @param {string} type
   * @param {{ column: number; row: number; value: CellValue; previousValue: CellValue; worksheet: Worksheet }} init
   */
  constructor(type, { column, row, value, previousValue, worksheet }) {
    super(type);

    /** @readonly */ this.column = column;
    /** @readonly */ this.row = row;
    /** @readonly */ this.value = value;
    /** @readonly */ this.previousValue = previousValue;
    /** @readonly */ this.worksheet = worksheet;
  }
}
