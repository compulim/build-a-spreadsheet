/** @typedef {import('../Worksheet.js').CellValue} CellValue */
/** @typedef {import('../Worksheet.js').Worksheet} Worksheet */

export class CellChangeEvent extends Event {
  /**
   * @param {string} type
   * @param {{ x: number; y: number; value: CellValue; previousValue: CellValue; worksheet: Worksheet }} init
   */
  constructor(type, { x, y, value, previousValue, worksheet }) {
    super(type);

    /** @readonly */ this.x = x;
    /** @readonly */ this.y = y;
    /** @readonly */ this.value = value;
    /** @readonly */ this.previousValue = previousValue;
    /** @readonly */ this.worksheet = worksheet;
  }
}
