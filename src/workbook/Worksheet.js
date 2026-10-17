import { parseCSV, stringifyCSV } from '../csv.js';
import { CellChangeEvent } from './private/CellChangeEvent.js';
import { COLUMN_COUNT, COLUMN_WIDTH, ROW_COUNT } from './constants.js';

/** @typedef {boolean | number | string} CellValue */
/** @typedef {{ top: boolean; right: boolean; bottom: boolean; left: boolean }} CellBorder */
/** @typedef {'left' | 'right' | 'top' | 'bottom' | 'outside' | 'all' | 'none'} BorderPreset */
/** @typedef {{ x: number; y: number; value: CellValue; previousValue: CellValue }} CellChange */
/** @typedef {CellValue | null | undefined} CellInput */
/** @typedef {(value: CellValue | undefined, x: number, y: number) => CellInput} CellUpdater */

/** @type {readonly (keyof CellBorder)[]} */
const BORDER_SIDES = ['top', 'right', 'bottom', 'left'];

/**
 * A single worksheet of 50 columns by 10,000 rows (by default). Values are stored sparsely.
 * All coordinates (`x` for column, `y` for row) are zero-based.
 *
 * Dispatches `cellchange` (`CellChangeEvent`) whenever the value of a cell changes, `borderchange` (with `x` and `y`)
 * whenever the border of a cell changes, and `columnwidthchange` (with `x`) whenever the width of a column changes.
 *
 * Keeps an undo history of cell changes. Changes made synchronously in the same task (e.g. a paste) are grouped into a
 * single undo step.
 */
export class Worksheet extends EventTarget {
  /** @type {(CellValue[] | undefined)[]} */
  #rows = [];

  /** Borders of cells, keyed by `y * columnCount + x`. @type {Map<number, CellBorder>} */
  #borders = new Map();

  /** @type {number[]} */
  #columnWidths;

  /** @type {CellChange[][]} */
  #undoStack = [];

  /** @type {CellChange[][]} */
  #redoStack = [];

  /** @type {CellChange[] | undefined} */
  #pending;

  #applyingHistory = false;

  /** @param {{ columnCount?: number; rowCount?: number; columnWidth?: number }} [options] */
  constructor({ columnCount = COLUMN_COUNT, rowCount = ROW_COUNT, columnWidth = COLUMN_WIDTH } = {}) {
    super();

    /** @readonly */ this.columnCount = columnCount;
    /** @readonly */ this.rowCount = rowCount;

    this.#columnWidths = new Array(columnCount).fill(columnWidth);
  }

  /**
   * @param {number} x
   * @param {number} y
   */
  #assertCell(x, y) {
    if (!Number.isInteger(x) || x < 0 || x >= this.columnCount) {
      throw new RangeError(`X must be an integer between 0 and ${this.columnCount - 1}.`);
    } else if (!Number.isInteger(y) || y < 0 || y >= this.rowCount) {
      throw new RangeError(`Y must be an integer between 0 and ${this.rowCount - 1}.`);
    }
  }

  /**
   * Returns the width of a column in pixels.
   *
   * @param {number} x
   * @returns {number}
   */
  getColumnWidth(x) {
    this.#assertCell(x, 0);

    return this.#columnWidths[x];
  }

  /**
   * Sets the width of a column in pixels. Dispatches `columnwidthchange` (with `x`) if changed.
   *
   * @param {number} x
   * @param {number} width
   */
  setColumnWidth(x, width) {
    this.#assertCell(x, 0);

    if (!Number.isFinite(width) || width < 0) {
      throw new RangeError('Width must be a non-negative finite number.');
    } else if (this.#columnWidths[x] === width) {
      return;
    }

    this.#columnWidths[x] = width;

    this.dispatchEvent(Object.assign(new Event('columnwidthchange'), { x }));
  }

  /**
   * Returns the value of a cell, empty cells return `undefined`.
   *
   * @param {number} x
   * @param {number} y
   * @returns {CellValue | undefined}
   */
  getValue(x, y) {
    this.#assertCell(x, y);

    return this.#rows[y]?.[x];
  }

  /**
   * Sets the value of a cell. Setting `''`, `null` or `undefined` clears the cell.
   *
   * @param {number} x
   * @param {number} y
   * @param {CellInput | CellUpdater} valueOrUpdater A value, or a function `(value, x, y) => nextValue`.
   */
  setValue(x, y, valueOrUpdater) {
    this.#assertCell(x, y);

    let value = typeof valueOrUpdater === 'function' ? valueOrUpdater(this.#rows[y]?.[x], x, y) : valueOrUpdater;

    if (value === null || value === undefined) {
      value = '';
    } else if (typeof value === 'number') {
      if (!Number.isFinite(value)) {
        throw new TypeError('Number value must be finite.');
      }
    } else if (typeof value !== 'boolean' && typeof value !== 'string') {
      throw new TypeError('Value must be a boolean, number or string.');
    }

    const cells = this.#rows[y];
    const previousValue = cells?.[x] ?? '';

    if (Object.is(previousValue, value)) {
      return;
    }

    if (value === '') {
      if (cells) {
        delete cells[x];

        if (!cells.some(() => true)) {
          delete this.#rows[y];
        }
      }
    } else {
      (cells || (this.#rows[y] = []))[x] = value;
    }

    this.#applyingHistory || this.#record({ x, y, value, previousValue });

    this.dispatchEvent(new CellChangeEvent('cellchange', { x, y, value, previousValue, worksheet: this }));
  }

  /** @param {CellChange} change */
  #record(change) {
    if (!this.#pending) {
      const pending = (this.#pending = []);

      this.#undoStack.push(pending);
      this.#redoStack = [];
      queueMicrotask(() => {
        this.#pending === pending && (this.#pending = undefined);
      });
    }

    this.#pending.push(change);
  }

  /** Size of the used area, from A1 to the last used row/column. */
  #getUsedSize() {
    let height = 0;
    let width = 0;

    this.#rows.forEach((cells, y) => {
      height = y + 1;
      cells.forEach((_, x) => {
        width = Math.max(width, x + 1);
      });
    });

    return { width, height };
  }

  /**
   * Returns values of a range as a 2D array (`values[y][x]`). Empty cells are returned as `''`.
   *
   * @param {number} x
   * @param {number} y
   * @param {number} width
   * @param {number} height
   * @returns {CellValue[][]}
   */
  getValues(x, y, width, height) {
    if (!Number.isInteger(width) || !Number.isInteger(height) || width < 0 || height < 0) {
      throw new RangeError('Width and height must be non-negative integers.');
    }

    if (!width || !height) {
      return [];
    }

    this.#assertCell(x, y);
    this.#assertCell(x + width - 1, y + height - 1);

    /** @type {CellValue[][]} */
    const values = [];

    for (let rowOffset = 0; rowOffset < height; rowOffset++) {
      const cells = this.#rows[y + rowOffset];
      const rowValues = new Array(width);

      for (let columnOffset = 0; columnOffset < width; columnOffset++) {
        rowValues[columnOffset] = cells?.[x + columnOffset] ?? '';
      }

      values.push(rowValues);
    }

    return values;
  }

  /**
   * Sets values from a 2D array (`values[y][x]`), starting at the given cell. Values outside of the worksheet are ignored.
   *
   * Alternatively, sets every cell of a range with an updater function `(value, x, y) => nextValue`, e.g.
   * `setValues(() => '', x, y, width, height)` clears the range.
   *
   * @param {CellInput[][] | CellUpdater} valuesOrUpdater
   * @param {number} [x=0]
   * @param {number} [y=0]
   * @param {number} [width] Required with an updater function.
   * @param {number} [height] Required with an updater function.
   */
  setValues(valuesOrUpdater, x = 0, y = 0, width, height) {
    this.#assertCell(x, y);

    if (typeof valuesOrUpdater === 'function') {
      if (!Number.isInteger(width) || !Number.isInteger(height) || width < 0 || height < 0) {
        throw new RangeError('Width and height must be non-negative integers.');
      }

      const right = Math.min(x + width, this.columnCount);
      const bottom = Math.min(y + height, this.rowCount);

      for (let cellY = y; cellY < bottom; cellY++) {
        for (let cellX = x; cellX < right; cellX++) {
          this.setValue(cellX, cellY, valuesOrUpdater);
        }
      }

      return;
    }

    const values = valuesOrUpdater;
    const rowCount = Math.min(values.length, this.rowCount - y);

    for (let rowOffset = 0; rowOffset < rowCount; rowOffset++) {
      const rowValues = values[rowOffset] || [];
      const columnCount = Math.min(rowValues.length, this.columnCount - x);

      for (let columnOffset = 0; columnOffset < columnCount; columnOffset++) {
        this.setValue(x + columnOffset, y + rowOffset, rowValues[columnOffset]);
      }
    }
  }

  /** Whether `undo()` has anything to revert. */
  get canUndo() {
    return this.#undoStack.length > 0;
  }

  /** Whether `redo()` has anything to re-apply. */
  get canRedo() {
    return this.#redoStack.length > 0;
  }

  /**
   * Reverts the last group of cell changes.
   *
   * @returns {boolean} `false` if there is nothing to undo.
   */
  undo() {
    return this.#applyHistory(this.#undoStack, this.#redoStack, true);
  }

  /**
   * Re-applies the last undone group of cell changes.
   *
   * @returns {boolean} `false` if there is nothing to redo.
   */
  redo() {
    return this.#applyHistory(this.#redoStack, this.#undoStack, false);
  }

  /**
   * @param {CellChange[][]} from
   * @param {CellChange[][]} to
   * @param {boolean} isUndo
   */
  #applyHistory(from, to, isUndo) {
    const changes = from.pop();

    if (!changes) {
      return false;
    }

    this.#pending = undefined;
    to.push(changes);
    this.#applyingHistory = true;

    try {
      for (const { x, y, value, previousValue } of isUndo ? [...changes].reverse() : changes) {
        this.setValue(x, y, isUndo ? previousValue : value);
      }
    } finally {
      this.#applyingHistory = false;
    }

    return true;
  }

  /**
   * Returns the borders of a cell.
   *
   * @param {number} x
   * @param {number} y
   * @returns {CellBorder}
   */
  getBorder(x, y) {
    this.#assertCell(x, y);

    return { top: false, right: false, bottom: false, left: false, ...this.#borders.get(y * this.columnCount + x) };
  }

  /**
   * Sets some or all borders of a cell. Dispatches `borderchange` (with `x` and `y`) if changed.
   *
   * @param {number} x
   * @param {number} y
   * @param {Partial<CellBorder>} border
   */
  setBorder(x, y, border) {
    const previous = this.getBorder(x, y);
    const next = { ...previous };

    for (const side of BORDER_SIDES) {
      side in border && (next[side] = !!border[side]);
    }

    if (BORDER_SIDES.every(side => next[side] === previous[side])) {
      return;
    }

    const key = y * this.columnCount + x;

    BORDER_SIDES.some(side => next[side]) ? this.#borders.set(key, next) : this.#borders.delete(key);

    this.dispatchEvent(Object.assign(new Event('borderchange'), { x, y }));
  }

  /**
   * Sets borders of a range:
   *
   * - `'left'`: left of the leftmost cells.
   * - `'right'`: right of the rightmost cells.
   * - `'top'`: top of the topmost cells.
   * - `'bottom'`: bottom of the bottommost cells.
   * - `'outside'`: surrounds the range.
   * - `'all'`: surrounds the range and every cell inside it.
   * - `'none'`: removes the surrounding and inside borders.
   *
   * @param {{ left: number; top: number; right: number; bottom: number }} range
   * @param {BorderPreset} preset
   */
  setRangeBorder({ left, top, right, bottom }, preset) {
    this.#assertCell(left, top);
    this.#assertCell(right, bottom);

    const outside = preset === 'outside';

    for (let y = top; y <= bottom; y++) {
      for (let x = left; x <= right; x++) {
        if (preset === 'all' || preset === 'none') {
          const set = preset === 'all';

          this.setBorder(x, y, { top: set, right: set, bottom: set, left: set });
        } else {
          /** @type {Partial<CellBorder>} */
          const border = {};

          (outside || preset === 'left') && x === left && (border.left = true);
          (outside || preset === 'right') && x === right && (border.right = true);
          (outside || preset === 'top') && y === top && (border.top = true);
          (outside || preset === 'bottom') && y === bottom && (border.bottom = true);

          this.setBorder(x, y, border);
        }
      }
    }

    // Edges shared with adjacent cells outside of the range are also part of the surrounding border.
    if (preset === 'none') {
      for (let y = top; y <= bottom; y++) {
        left > 0 && this.setBorder(left - 1, y, { right: false });
        right < this.columnCount - 1 && this.setBorder(right + 1, y, { left: false });
      }

      for (let x = left; x <= right; x++) {
        top > 0 && this.setBorder(x, top - 1, { bottom: false });
        bottom < this.rowCount - 1 && this.setBorder(x, bottom + 1, { top: false });
      }
    }
  }

  /**
   * Serializes the worksheet into CSV.
   *
   * @returns {string}
   */
  toCSV() {
    const { width, height } = this.#getUsedSize();

    return stringifyCSV(this.getValues(0, 0, width, height));
  }

  /**
   * Creates a worksheet from CSV.
   *
   * @param {string} text
   * @returns {Worksheet}
   */
  static fromCSV(text) {
    const worksheet = new Worksheet();

    worksheet.setValues(parseCSV(text));
    // Loading is not an undoable change.
    worksheet.#undoStack = [];
    worksheet.#pending = undefined;

    return worksheet;
  }
}
