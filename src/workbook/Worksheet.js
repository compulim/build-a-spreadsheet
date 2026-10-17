import { parseCSV, stringifyCSV } from '../csv.js';
import { CellChangeEvent } from './private/CellChangeEvent.js';
import { COLUMN_COUNT, COLUMN_WIDTH, ROW_COUNT } from './constants.js';

/** @typedef {boolean | number | string} CellValue */
/** @typedef {{ top: boolean; right: boolean; bottom: boolean; left: boolean }} CellBorder */
/** @typedef {'left' | 'right' | 'top' | 'bottom' | 'outside' | 'inside'} BorderEdges */
/** @typedef {{ x: number; y: number; value: CellValue; previousValue: CellValue }} CellValueChange */
/** @typedef {{ x: number; y: number; formatting: CellFormatting; previousFormatting: CellFormatting }} CellFormattingChange */
/** @typedef {CellValueChange | CellFormattingChange} CellChange */
/** @typedef {CellValue | null | undefined} CellInput */
/** @typedef {(value: CellValue | undefined, x: number, y: number) => CellInput} CellUpdater */
/** @typedef {'left' | 'center' | 'right'} CellAlign */
/**
 * `align` is `undefined` when not set: text is aligned left, numbers right and booleans center.
 * `highlight` is the name of a highlight (background/foreground color combination), `undefined` when not set.
 * @typedef {{ bold: boolean; italic: boolean; underline: boolean; strikethrough: boolean; align: CellAlign | undefined; highlight: string | undefined }} CellFormatting
 */
/** @typedef {Partial<CellFormatting> | null | undefined} FormattingInput */
/** @typedef {(formatting: CellFormatting, x: number, y: number) => FormattingInput} FormattingUpdater */

/** @type {Readonly<CellFormatting>} */
const DEFAULT_FORMATTING = Object.freeze({
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  align: undefined,
  highlight: undefined
});

/**
 * A single worksheet of 50 columns by 10,000 rows (by default). Values are stored sparsely.
 * All coordinates (`x` for column, `y` for row) are zero-based.
 *
 * Dispatches `cellchange` (`CellChangeEvent`) whenever the value of a cell changes, `borderchange` (with `x` and `y`)
 * whenever the border of a cell changes, `formattingchange` (with `x` and `y`) whenever the formatting of a cell changes,
 * and `columnwidthchange` (with `x`) whenever the width of a column changes.
 *
 * Keeps an undo history of cell value and formatting changes. Changes made synchronously in the same task (e.g. a paste) are grouped into a
 * single undo step.
 */
export class Worksheet extends EventTarget {
  /** @type {(CellValue[] | undefined)[]} */
  #rows = [];

  /**
   * Left and top borders of cells, keyed by `y * columnCount + x`. The right and bottom borders of a cell are the left
   * border of the cell on the right and the top border of the cell below.
   * @type {Map<number, { top: boolean; left: boolean }>}
   */
  #borders = new Map();

  /**
   * Formatting of cells, keyed by `y * columnCount + x`. Unformatted cells are not stored.
   * @type {Map<number, CellFormatting>}
   */
  #formattings = new Map();

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
   * Sets the value of a cell. Strings are trimmed. Setting an empty or whitespace-only string, `null` or `undefined` clears the cell.
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
    } else if (typeof value === 'string') {
      value = value.trim();
    } else if (typeof value !== 'boolean') {
      throw new TypeError('Value must be a boolean, number or string.');
    }

    const cells = this.#rows[y];
    const previousValue = cells?.[x] ?? '';

    if (Object.is(previousValue, value)) {
      return;
    }

    if (value === '') {
      // Empty cells never keep formatting. Removed (and recorded) before the value, so undo restores the value first.
      this.#setFormatting(x, y, null);

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
   * Reverts the last group of cell value and formatting changes.
   *
   * @returns {boolean} `false` if there is nothing to undo.
   */
  undo() {
    return this.#applyHistory(this.#undoStack, this.#redoStack, true);
  }

  /**
   * Re-applies the last undone group of cell value and formatting changes.
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
      for (const change of isUndo ? [...changes].reverse() : changes) {
        if ('formatting' in change) {
          this.#setFormatting(change.x, change.y, isUndo ? change.previousFormatting : change.formatting);
        } else {
          this.setValue(change.x, change.y, isUndo ? change.previousValue : change.value);
        }
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

    const own = this.#borders.get(y * this.columnCount + x);

    return {
      top: !!own?.top,
      right: x + 1 < this.columnCount && !!this.#borders.get(y * this.columnCount + x + 1)?.left,
      bottom: y + 1 < this.rowCount && !!this.#borders.get((y + 1) * this.columnCount + x)?.top,
      left: !!own?.left
    };
  }

  /**
   * @param {number} x
   * @param {number} y
   * @param {'left' | 'top'} side
   * @param {boolean} value
   */
  #setStoredEdge(x, y, side, value) {
    const key = y * this.columnCount + x;
    const previous = this.#borders.get(key) || { top: false, left: false };

    if (previous[side] === value) {
      return;
    }

    const next = { ...previous, [side]: value };

    next.top || next.left ? this.#borders.set(key, next) : this.#borders.delete(key);

    this.dispatchEvent(Object.assign(new Event('borderchange'), { x, y }));
  }

  /**
   * Sets or unsets borders of a range:
   *
   * - `'left'`: left of the leftmost cells.
   * - `'right'`: right of the rightmost cells.
   * - `'top'`: top of the topmost cells.
   * - `'bottom'`: bottom of the bottommost cells.
   * - `'outside'`: all four sides of the range.
   * - `'inside'`: between cells inside the range.
   *
   * Omit `x` (and `width`) to span whole rows, omit `y` (and `height`) to span whole columns. `width` and `height`
   * default to 1 when `x` or `y` is set. For example, `{ x: 2 }` is column C and `{ y: 4 }` is row 5.
   *
   * The right and bottom borders are stored as the left border of the cell on the right and the top border of the
   * cell below, and are ignored on the last column and row.
   * Dispatches `borderchange` (with `x` and `y` of the cell storing the edge) for each stored edge changed.
   *
   * `edges` is a `Set` of edges to apply together, for example `new Set(['top', 'left'])`. `'outside'` is a shortcut
   * of `'left'`, `'right'`, `'top'` and `'bottom'`.
   *
   * @param {{ x?: number; y?: number; width?: number; height?: number }} range
   * @param {ReadonlySet<BorderEdges>} edges
   * @param {boolean} [value] `false` to unset the borders.
   */
  setBorder(range, edges, value = true) {
    const left = range.x ?? 0;
    const top = range.y ?? 0;
    const right = left + (range.width ?? (range.x === undefined ? this.columnCount : 1)) - 1;
    const bottom = top + (range.height ?? (range.y === undefined ? this.rowCount : 1)) - 1;

    this.#assertCell(left, top);
    this.#assertCell(right, bottom);

    const set = new Set(edges);

    if (set.delete('outside')) {
      set.add('left').add('right').add('top').add('bottom');
    }

    for (let y = top; y <= bottom; y++) {
      for (let x = left; x <= right; x++) {
        if (set.has('inside')) {
          x > left && this.#setStoredEdge(x, y, 'left', value);
          y > top && this.#setStoredEdge(x, y, 'top', value);
        }

        set.has('left') && x === left && this.#setStoredEdge(x, y, 'left', value);
        set.has('top') && y === top && this.#setStoredEdge(x, y, 'top', value);
        set.has('right') && x === right && x + 1 < this.columnCount && this.#setStoredEdge(x + 1, y, 'left', value);
        set.has('bottom') && y === bottom && y + 1 < this.rowCount && this.#setStoredEdge(x, y + 1, 'top', value);
      }
    }
  }

  /**
   * Sets the formatting of a cell. Missing properties are set to default. Setting `null`, `undefined` or the default
   * formatting removes the formatting. Empty cells cannot be formatted and are left as-is. Dispatches
   * `formattingchange` (with `x` and `y`) if changed.
   *
   * @param {number} x
   * @param {number} y
   * @param {FormattingInput | FormattingUpdater} formattingOrUpdater A formatting, or a function `(formatting, x, y) => nextFormatting`.
   */
  #setFormatting(x, y, formattingOrUpdater) {
    if ((this.#rows[y]?.[x] ?? '') === '') {
      return;
    }

    const key = y * this.columnCount + x;
    const input =
      typeof formattingOrUpdater === 'function'
        ? formattingOrUpdater({ ...(this.#formattings.get(key) || DEFAULT_FORMATTING) }, x, y)
        : formattingOrUpdater;

    if (input !== null && input !== undefined && typeof input !== 'object') {
      throw new TypeError('Formatting must be an object.');
    }

    const align = input?.align ?? undefined;

    if (align !== undefined && align !== 'left' && align !== 'center' && align !== 'right') {
      throw new TypeError('Align must be "left", "center" or "right".');
    }

    const highlight = input?.highlight || undefined;

    if (highlight !== undefined && typeof highlight !== 'string') {
      throw new TypeError('Highlight must be a string.');
    }

    /** @type {CellFormatting} */
    const formatting = {
      bold: !!input?.bold,
      italic: !!input?.italic,
      underline: !!input?.underline,
      strikethrough: !!input?.strikethrough,
      align,
      highlight
    };

    const previous = this.#formattings.get(key) || DEFAULT_FORMATTING;
    const isEqual = (/** @type {CellFormatting} */ a, /** @type {CellFormatting} */ b) =>
      a.bold === b.bold &&
      a.italic === b.italic &&
      a.underline === b.underline &&
      a.strikethrough === b.strikethrough &&
      a.align === b.align &&
      a.highlight === b.highlight;

    if (isEqual(previous, formatting)) {
      return;
    }

    isEqual(formatting, DEFAULT_FORMATTING) ? this.#formattings.delete(key) : this.#formattings.set(key, formatting);

    this.#applyingHistory || this.#record({ x, y, formatting, previousFormatting: { ...previous } });

    this.dispatchEvent(Object.assign(new Event('formattingchange'), { x, y }));
  }

  /**
   * Returns formattings of a range as a 2D array (`formattings[y][x]`). Unformatted cells return the default
   * formatting (no styles, `align` of `undefined`).
   *
   * @param {number} x
   * @param {number} y
   * @param {number} width
   * @param {number} height
   * @returns {CellFormatting[][]}
   */
  getFormattings(x, y, width, height) {
    if (!Number.isInteger(width) || !Number.isInteger(height) || width < 0 || height < 0) {
      throw new RangeError('Width and height must be non-negative integers.');
    }

    if (!width || !height) {
      return [];
    }

    this.#assertCell(x, y);
    this.#assertCell(x + width - 1, y + height - 1);

    /** @type {CellFormatting[][]} */
    const formattings = [];

    for (let rowOffset = 0; rowOffset < height; rowOffset++) {
      const row = [];

      for (let columnOffset = 0; columnOffset < width; columnOffset++) {
        row.push({
          ...(this.#formattings.get((y + rowOffset) * this.columnCount + x + columnOffset) || DEFAULT_FORMATTING)
        });
      }

      formattings.push(row);
    }

    return formattings;
  }

  /**
   * Sets formattings from a 2D array (`formattings[y][x]`), starting at the given cell. Formattings outside of the
   * worksheet are ignored.
   *
   * Alternatively, sets every cell of a range with an updater function `(formatting, x, y) => nextFormatting`, e.g.
   * `setFormattings(() => null, x, y, width, height)` removes formatting of the range.
   *
   * @param {FormattingInput[][] | FormattingUpdater} formattingsOrUpdater
   * @param {number} [x=0]
   * @param {number} [y=0]
   * @param {number} [width] Required with an updater function.
   * @param {number} [height] Required with an updater function.
   */
  setFormattings(formattingsOrUpdater, x = 0, y = 0, width, height) {
    this.#assertCell(x, y);

    if (typeof formattingsOrUpdater === 'function') {
      if (!Number.isInteger(width) || !Number.isInteger(height) || width < 0 || height < 0) {
        throw new RangeError('Width and height must be non-negative integers.');
      }

      const right = Math.min(x + width, this.columnCount);
      const bottom = Math.min(y + height, this.rowCount);

      for (let cellY = y; cellY < bottom; cellY++) {
        for (let cellX = x; cellX < right; cellX++) {
          this.#setFormatting(cellX, cellY, formattingsOrUpdater);
        }
      }

      return;
    }

    const formattings = formattingsOrUpdater;
    const rowCount = Math.min(formattings.length, this.rowCount - y);

    for (let rowOffset = 0; rowOffset < rowCount; rowOffset++) {
      const row = formattings[rowOffset] || [];
      const columnCount = Math.min(row.length, this.columnCount - x);

      for (let columnOffset = 0; columnOffset < columnCount; columnOffset++) {
        this.#setFormatting(x + columnOffset, y + rowOffset, row[columnOffset]);
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
