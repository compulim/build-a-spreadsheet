import { parseCSV, stringifyCSV } from './csv.js';
import { createZip, readZip } from './zip.js';

/** @typedef {boolean | number | string} CellValue */

export const COLUMN_COUNT = 50;
export const ROW_COUNT = 10_000;
export const MAX_WORKSHEETS = 50;
export const MAX_WORKSHEET_NAME_LENGTH = 31;

// Same restrictions as Excel, also make sure the name is a valid file name inside the ZIP file.
const INVALID_NAME_PATTERN = /[\\/?*:[\]\u0000-\u001f]/;

/** Worksheet -> Workbook it belongs to. */
const owners = new WeakMap();

/**
 * Validates a worksheet name. Returns an error message, or `undefined` if the name is valid.
 *
 * @param {string} name
 * @returns {string | undefined}
 */
export function validateWorksheetName(name) {
  if (typeof name !== 'string' || !name.trim()) {
    return 'Sheet name cannot be empty.';
  } else if (name.length > MAX_WORKSHEET_NAME_LENGTH) {
    return `Sheet name cannot be longer than ${MAX_WORKSHEET_NAME_LENGTH} characters.`;
  } else if (INVALID_NAME_PATTERN.test(name)) {
    return 'Sheet name cannot contain any of these characters: \\ / ? * : [ ]';
  } else if (name.startsWith("'") || name.endsWith("'")) {
    return 'Sheet name cannot begin or end with an apostrophe.';
  }
}

/**
 * Turns any string into a valid worksheet name.
 *
 * @param {string} name
 * @returns {string}
 */
export function sanitizeWorksheetName(name) {
  const sanitized = String(name)
    .replace(new RegExp(INVALID_NAME_PATTERN, 'g'), '_')
    .replace(/^'+|'+$/g, '')
    .trim()
    .slice(0, MAX_WORKSHEET_NAME_LENGTH)
    .trim();

  return sanitized || 'Sheet';
}

/**
 * Converts a zero-based column index into its letter name, e.g. 0 -> "A", 26 -> "AA".
 *
 * @param {number} column
 * @returns {string}
 */
export function getColumnName(column) {
  let name = '';

  for (let index = column + 1; index > 0; index = Math.floor((index - 1) / 26)) {
    name = String.fromCharCode(65 + ((index - 1) % 26)) + name;
  }

  return name;
}

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

/**
 * A single worksheet of 50 columns by 10,000 rows (by default). Values are stored sparsely.
 *
 * Dispatches `cellchange` (`CellChangeEvent`) whenever the value of a cell changes.
 */
export class Worksheet extends EventTarget {
  /** @type {string} */
  #name;

  /** @type {(CellValue[] | undefined)[]} */
  #rows = [];

  /**
   * @param {string} [name]
   * @param {{ columnCount?: number; rowCount?: number }} [options]
   */
  constructor(name = 'Sheet1', { columnCount = COLUMN_COUNT, rowCount = ROW_COUNT } = {}) {
    super();

    const error = validateWorksheetName(name);

    if (error) {
      throw new TypeError(error);
    }

    this.#name = name;

    /** @readonly */ this.columnCount = columnCount;
    /** @readonly */ this.rowCount = rowCount;
  }

  get name() {
    return this.#name;
  }

  set name(name) {
    if (name === this.#name) {
      return;
    }

    const error = validateWorksheetName(name);

    if (error) {
      throw new TypeError(error);
    }

    const workbook = owners.get(this);

    if (workbook && workbook.getWorksheet(name) && workbook.getWorksheet(name) !== this) {
      throw new TypeError(`A sheet named "${name}" already exists.`);
    }

    const previousName = this.#name;

    this.#name = name;

    workbook?.dispatchEvent(
      new WorksheetEvent('worksheetrename', { worksheet: this, index: workbook.indexOf(this), previousName })
    );
  }

  /** The workbook this worksheet belongs to, if any. */
  get workbook() {
    return owners.get(this);
  }

  /**
   * @param {number} column
   * @param {number} row
   */
  #assertCell(column, row) {
    if (!Number.isInteger(column) || column < 0 || column >= this.columnCount) {
      throw new RangeError(`Column must be an integer between 0 and ${this.columnCount - 1}.`);
    } else if (!Number.isInteger(row) || row < 0 || row >= this.rowCount) {
      throw new RangeError(`Row must be an integer between 0 and ${this.rowCount - 1}.`);
    }
  }

  /**
   * Returns the value of a cell, empty cells return `''`.
   *
   * @param {number} column Zero-based column index.
   * @param {number} row Zero-based row index.
   * @returns {CellValue}
   */
  getValue(column, row) {
    this.#assertCell(column, row);

    return this.#rows[row]?.[column] ?? '';
  }

  /**
   * Sets the value of a cell. Setting `''`, `null` or `undefined` clears the cell.
   *
   * @param {number} column Zero-based column index.
   * @param {number} row Zero-based row index.
   * @param {CellValue} value
   */
  setValue(column, row, value) {
    this.#assertCell(column, row);

    if (value === null || value === undefined) {
      value = '';
    } else if (typeof value === 'number') {
      if (!Number.isFinite(value)) {
        throw new TypeError('Number value must be finite.');
      }
    } else if (typeof value !== 'boolean' && typeof value !== 'string') {
      throw new TypeError('Value must be a boolean, number or string.');
    }

    const cells = this.#rows[row];
    const previousValue = cells?.[column] ?? '';

    if (Object.is(previousValue, value)) {
      return;
    }

    if (value === '') {
      if (cells) {
        delete cells[column];

        if (!cells.some(() => true)) {
          delete this.#rows[row];
        }
      }
    } else {
      (cells || (this.#rows[row] = []))[column] = value;
    }

    this.dispatchEvent(new CellChangeEvent('cellchange', { column, row, value, previousValue, worksheet: this }));
  }

  /**
   * Returns all values as a 2D array (`values[row][column]`), starting from cell A1 up to the last used row/column.
   * Empty cells are returned as `''`. An empty worksheet returns `[]`.
   *
   * @returns {CellValue[][]}
   */
  getValues() {
    let lastRow = -1;
    let lastColumn = -1;

    this.#rows.forEach((cells, row) => {
      lastRow = row;
      cells.forEach((_, column) => {
        lastColumn = Math.max(lastColumn, column);
      });
    });

    /** @type {CellValue[][]} */
    const values = [];

    for (let row = 0; row <= lastRow; row++) {
      const cells = this.#rows[row];
      const rowValues = new Array(lastColumn + 1);

      for (let column = 0; column <= lastColumn; column++) {
        rowValues[column] = cells?.[column] ?? '';
      }

      values.push(rowValues);
    }

    return values;
  }

  /**
   * Sets values from a 2D array (`values[row][column]`), starting at the given cell.
   * Values outside of the worksheet are ignored.
   *
   * @param {CellValue[][]} values
   * @param {number} [column=0]
   * @param {number} [row=0]
   */
  setValues(values, column = 0, row = 0) {
    this.#assertCell(column, row);

    const rowCount = Math.min(values.length, this.rowCount - row);

    for (let rowOffset = 0; rowOffset < rowCount; rowOffset++) {
      const rowValues = values[rowOffset] || [];
      const columnCount = Math.min(rowValues.length, this.columnCount - column);

      for (let columnOffset = 0; columnOffset < columnCount; columnOffset++) {
        this.setValue(column + columnOffset, row + rowOffset, rowValues[columnOffset]);
      }
    }
  }

  /** Clears all cells. */
  clear() {
    this.#rows.forEach((cells, row) => {
      cells.forEach((_, column) => this.setValue(column, row, ''));
    });
  }

  /**
   * Serializes the worksheet into CSV.
   *
   * @returns {string}
   */
  toCSV() {
    return stringifyCSV(this.getValues());
  }

  /**
   * Replaces the content of this worksheet with CSV. Values outside of the worksheet are ignored.
   *
   * @param {string} text
   */
  loadCSV(text) {
    this.clear();
    this.setValues(parseCSV(text));
  }

  /**
   * Creates a worksheet from CSV.
   *
   * @param {string} text
   * @param {string} [name]
   * @returns {Worksheet}
   */
  static fromCSV(text, name) {
    const worksheet = new Worksheet(name);

    worksheet.setValues(parseCSV(text));

    return worksheet;
  }
}

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
