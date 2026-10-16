import { createZip, readZip } from '../zip.js';
import { CellChangeEvent } from './private/CellChangeEvent.js';
import { MAX_WORKSHEET_NAME_LENGTH, MAX_WORKSHEETS } from './constants.js';
import { owners } from './private/owners.js';
import { Worksheet } from './Worksheet.js';
import { sanitizeWorksheetName } from './private/worksheetName.js';
import { WorksheetEvent } from './private/WorksheetEvent.js';

/**
 * A workbook with up to 50 worksheets.
 *
 * Re-dispatches `cellchange` (`CellChangeEvent`) from all of its worksheets, and dispatches
 * `worksheetadd`, `worksheetremove` and `worksheetrename` (`WorksheetEvent`).
 */
export class Workbook extends EventTarget {
  /** @type {Worksheet[]} */
  #worksheets = [];

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

    return this.#worksheets.find(({ name }) => name.toLowerCase() === lowerName);
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
   * Adds a worksheet.
   *
   * @param {string | Worksheet} [nameOrWorksheet] Name of the new worksheet, or an existing worksheet. Default to an unused name.
   * @param {number} [index] Position of the new worksheet, default to the end.
   * @returns {Worksheet}
   */
  addWorksheet(nameOrWorksheet = this.getUniqueWorksheetName(), index = this.#worksheets.length) {
    if (this.#worksheets.length >= MAX_WORKSHEETS) {
      throw new RangeError(`A workbook cannot have more than ${MAX_WORKSHEETS} sheets.`);
    }

    const worksheet = typeof nameOrWorksheet === 'string' ? new Worksheet(nameOrWorksheet) : nameOrWorksheet;

    if (!(worksheet instanceof Worksheet)) {
      throw new TypeError('Argument must be a string or Worksheet.');
    } else if (owners.has(worksheet)) {
      throw new TypeError('Worksheet already belongs to a workbook.');
    } else if (this.getWorksheet(worksheet.name)) {
      throw new TypeError(`A sheet named "${worksheet.name}" already exists.`);
    }

    index = Math.max(0, Math.min(index, this.#worksheets.length));

    owners.set(worksheet, this);
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
    owners.delete(worksheet);
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
        name: `${worksheet.name}.csv`,
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

      workbook.addWorksheet(
        Worksheet.fromCSV(decoder.decode(content), workbook.getUniqueWorksheetName(fileName.replace(/\.csv$/i, '')))
      );
    }

    return workbook;
  }
}
