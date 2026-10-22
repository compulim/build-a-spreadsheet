/** @typedef {import('./Worksheet.js').Worksheet} Worksheet */
/** @typedef {import('./Worksheet.js').WorksheetChange} WorksheetChange */

/** Dispatched once per `Worksheet#edit()` transaction, with all of its changes in order. */
export class WorksheetChangeEvent extends Event {
  /**
   * @param {string} type
   * @param {{ changes: readonly WorksheetChange[]; load: boolean; worksheet: Worksheet }} init
   */
  constructor(type, { changes, load, worksheet }) {
    super(type);

    /** @readonly */ this.changes = changes;
    /** `true` if the transaction loaded CSV, which is not meant to be undoable. @readonly */ this.load = load;
    /** @readonly */ this.worksheet = worksheet;
  }
}
