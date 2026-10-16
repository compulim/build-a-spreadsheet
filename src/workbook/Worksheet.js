import { parseCSV, stringifyCSV } from '../csv.js';
import { CellChangeEvent } from './CellChangeEvent.js';
import { COLUMN_COUNT, ROW_COUNT } from './constants.js';
import { owners } from './owners.js';
import { validateWorksheetName } from './worksheetName.js';
import { WorksheetEvent } from './WorksheetEvent.js';

/** @typedef {boolean | number | string} CellValue */

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
