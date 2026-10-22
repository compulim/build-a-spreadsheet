import { createZip, readZip } from '../zip.js';
import { CellChangeEvent } from './private/CellChangeEvent.js';
import { MAX_WORKSHEET_NAME_LENGTH, MAX_WORKSHEETS } from './constants.js';
import { Worksheet } from './Worksheet.js';
import { sanitizeWorksheetName, validateWorksheetName } from './private/worksheetName.js';
import { WorksheetEvent } from './private/WorksheetEvent.js';

/**
 * A workbook with up to 50 worksheets.
 *
 * Worksheets are created by the workbook. In a browser where `<compulim-worksheet>` is defined, each worksheet is owned
 * by a newly instantiated `<compulim-worksheet>` element, see `getWorksheetElement()`.
 *
 * Re-dispatches `cellchange` (`CellChangeEvent`) from all of its worksheets, and dispatches
 * `worksheetadd`, `worksheetremove` and `worksheetrename` (`WorksheetEvent`).
 */
export class Workbook extends EventTarget {
  /** @type {Worksheet[]} */
  #worksheets = [];

  /** @type {Map<Worksheet, string>} */
  #names = new Map();

  /** @type {Map<Worksheet, HTMLElement & { readonly worksheet: Worksheet }>} */
  #elements = new Map();

  /** @param {Event} event */
  #handleCellChange = event => {
    this.dispatchEvent(new CellChangeEvent('cellchange', /** @type {CellChangeEvent} */ (event)));
  };

  /** @returns {Worksheet[]} */
  get worksheets() {
    return [...this.#worksheets];
  }

  /**
   * @param {Worksheet} worksheet
   * @returns {number}
   */
  indexOf(worksheet) {
    return this.#worksheets.indexOf(worksheet);
  }

  /**
   * Finds a worksheet by name (case-insensitive) or by index.
   *
   * @param {string | number} nameOrIndex
   * @returns {Worksheet | undefined}
   */
  getWorksheet(nameOrIndex) {
    if (typeof nameOrIndex === 'number') {
      return this.#worksheets[nameOrIndex];
    }

    const lowerName = nameOrIndex.toLowerCase();

    return this.#worksheets.find(worksheet => this.#names.get(worksheet)?.toLowerCase() === lowerName);
  }

  /**
   * Returns the name of a worksheet in this workbook, or `undefined` if it does not belong to this workbook.
   *
   * @param {Worksheet} worksheet
   * @returns {string | undefined}
   */
  getWorksheetName(worksheet) {
    return this.#names.get(worksheet);
  }

  /**
   * @param {string} name
   * @param {Worksheet} [except]
   */
  #assertName(name, except) {
    const error = validateWorksheetName(name);

    if (error) {
      throw new TypeError(error);
    }

    const existing = this.getWorksheet(name);

    if (existing && existing !== except) {
      throw new TypeError(`A sheet named "${name}" already exists.`);
    }
  }

  /**
   * Renames a worksheet. Dispatches `worksheetrename` if changed.
   *
   * @param {Worksheet} worksheet
   * @param {string} name
   */
  renameWorksheet(worksheet, name) {
    const previousName = this.#names.get(worksheet);

    if (previousName === undefined) {
      throw new RangeError('Worksheet not found.');
    } else if (name === previousName) {
      return;
    }

    this.#assertName(name, worksheet);
    this.#names.set(worksheet, name);

    this.dispatchEvent(
      new WorksheetEvent('worksheetrename', { worksheet, index: this.indexOf(worksheet), previousName })
    );
  }

  /**
   * Returns an unused name, based on `base`.
   *
   * @param {string} [base]
   * @returns {string}
   */
  getUniqueWorksheetName(base) {
    if (base === undefined) {
      for (let index = this.#worksheets.length + 1; ; index++) {
        if (!this.getWorksheet(`Sheet${index}`)) {
          return `Sheet${index}`;
        }
      }
    }

    const name = sanitizeWorksheetName(base);

    if (!this.getWorksheet(name)) {
      return name;
    }

    for (let index = 2; ; index++) {
      const suffix = ` (${index})`;
      const candidate = name.slice(0, MAX_WORKSHEET_NAME_LENGTH - suffix.length) + suffix;

      if (!this.getWorksheet(candidate)) {
        return candidate;
      }
    }
  }

  /**
   * Returns the `<compulim-worksheet>` element owning a worksheet of this workbook, or `undefined` if not found or
   * `<compulim-worksheet>` is not defined (e.g. outside of a browser).
   *
   * @param {Worksheet} worksheet
   * @returns {(HTMLElement & { readonly worksheet: Worksheet }) | undefined}
   */
  getWorksheetElement(worksheet) {
    return this.#elements.get(worksheet);
  }

  /**
   * Adds a new worksheet, owned by a newly instantiated `<compulim-worksheet>` if it is defined.
   *
   * @param {string} [name] Name of the worksheet, default to an unused name.
   * @param {number} [index] Position of the worksheet, default to the end.
   * @returns {Worksheet}
   */
  addWorksheet(name = this.getUniqueWorksheetName(), index = this.#worksheets.length) {
    if (this.#worksheets.length >= MAX_WORKSHEETS) {
      throw new RangeError(`A workbook cannot have more than ${MAX_WORKSHEETS} sheets.`);
    }

    this.#assertName(name);

    index = Math.max(0, Math.min(index, this.#worksheets.length));

    /** @type {(HTMLElement & { readonly worksheet: Worksheet }) | undefined} */
    const element = globalThis.customElements?.get('compulim-worksheet') && document.createElement('compulim-worksheet');
    const worksheet = element ? element.worksheet : new Worksheet();

    element && this.#elements.set(worksheet, element);
    this.#names.set(worksheet, name);
    this.#worksheets.splice(index, 0, worksheet);
    worksheet.addEventListener('cellchange', this.#handleCellChange);

    this.dispatchEvent(new WorksheetEvent('worksheetadd', { worksheet, index }));

    return worksheet;
  }

  /**
   * Removes a worksheet.
   *
   * @param {Worksheet | number} worksheetOrIndex
   * @returns {Worksheet}
   */
  removeWorksheet(worksheetOrIndex) {
    const index =
      typeof worksheetOrIndex === 'number' ? worksheetOrIndex : this.#worksheets.indexOf(worksheetOrIndex);
    const worksheet = this.#worksheets[index];

    if (!worksheet) {
      throw new RangeError('Worksheet not found.');
    }

    this.#worksheets.splice(index, 1);
    this.#elements.get(worksheet)?.remove();
    this.#elements.delete(worksheet);
    this.#names.delete(worksheet);
    worksheet.removeEventListener('cellchange', this.#handleCellChange);

    this.dispatchEvent(new WorksheetEvent('worksheetremove', { worksheet, index }));

    return worksheet;
  }

  /**
   * Saves all worksheets as CSV files (named `<sheet name>.csv`) into a ZIP file.
   *
   * @returns {Promise<Uint8Array>}
   */
  async toZip() {
    const encoder = new TextEncoder();

    return createZip(
      this.#worksheets.map(worksheet => ({
        name: `${this.#names.get(worksheet)}.csv`,
        // Byte order mark so Excel recognizes the CSV file as UTF-8.
        data: encoder.encode(`\ufeff${worksheet.toCSV()}`)
      }))
    );
  }

  /**
   * Loads a workbook from a ZIP file containing CSV files, one worksheet per CSV file (in the order stored in the ZIP).
   *
   * @param {ArrayBuffer | Uint8Array} data
   * @returns {Promise<Workbook>}
   */
  static async fromZip(data) {
    const decoder = new TextDecoder();
    const workbook = new Workbook();

    for (const { name, data: content } of await readZip(data)) {
      const fileName = name.split('/').pop() || '';

      if (!/\.csv$/i.test(fileName) || name.startsWith('__MACOSX/') || fileName.startsWith('._')) {
        continue;
      }

      if (workbook.#worksheets.length >= MAX_WORKSHEETS) {
        break;
      }

      workbook
        .addWorksheet(workbook.getUniqueWorksheetName(fileName.replace(/\.csv$/i, '')))
        .loadCSV(decoder.decode(content));
    }

    return workbook;
  }
}
