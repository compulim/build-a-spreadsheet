/** @typedef {import('./Worksheet.ts').Worksheet} Worksheet */

export class WorksheetEvent extends Event {
  /**
   * @param {string} type
   * @param {{ worksheet: Worksheet; index: number; previousName?: string }} init
   */
  constructor(type, { worksheet, index, previousName }) {
    super(type);

    /** @readonly */ this.worksheet = worksheet;
    /** @readonly */ this.index = index;
    /** @readonly */ this.previousName = previousName;
  }
}
