import { formatValue, parseCSV, stringifyCSV } from '../csv.js';
import { WorksheetChangeEvent } from './WorksheetChangeEvent.js';
import { COLUMN_COUNT, COLUMN_WIDTH, ROW_COUNT } from './constants.js';

/** @typedef {boolean | number | string} CellValue */
/** @typedef {{ top: boolean; right: boolean; bottom: boolean; left: boolean }} CellBorder */
/** @typedef {'left' | 'right' | 'top' | 'bottom' | 'outside' | 'inside'} BorderEdges */
/** @typedef {{ type: 'value'; x: number; y: number; value: CellValue; previousValue: CellValue }} CellValueChange */
/** @typedef {{ type: 'formatting'; x: number; y: number; formatting: CellFormatting; previousFormatting: CellFormatting }} CellFormattingChange */
/** @typedef {{ width: number; height: number }} CellMergeSize */
/** @typedef {{ type: 'merge'; x: number; y: number; merge: CellMergeSize | undefined; previousMerge: CellMergeSize | undefined }} CellMergeChange */
/** @typedef {{ x: number; y: number; width: number; height: number }} CellMerge */
/** @typedef {{ type: 'border'; x: number; y: number; edge: 'left' | 'top'; border: boolean; previousBorder: boolean }} CellBorderChange */
/** @typedef {{ type: 'columnwidth'; x: number; width: number; previousWidth: number }} ColumnWidthChange */
/** @typedef {CellValueChange | CellFormattingChange | CellMergeChange | CellBorderChange} CellChange */
/** @typedef {CellChange | ColumnWidthChange} WorksheetChange */
/** @typedef {CellValue | null | undefined} CellInput */
/** @typedef {(value: CellValue | undefined, x: number, y: number) => CellInput} CellUpdater */
/** @typedef {'left' | 'center' | 'right'} CellAlign */
/** @typedef {'currency' | 'date' | 'datetime' | 'duration' | 'percent'} CellNumberFormat */
/**
 * Editing functions, only usable inside the `edit()` callback.
 * @typedef {{
 *   setColumnWidth(x: number, width: number): void;
 *   setValue(x: number, y: number, valueOrUpdater: CellInput | CellUpdater): void;
 *   setValues(valuesOrUpdater: CellInput[][] | CellUpdater, x?: number, y?: number, width?: number, height?: number): void;
 *   setFormattings(formattingsOrUpdater: FormattingInput[][] | FormattingUpdater, x?: number, y?: number, width?: number, height?: number): void;
 *   setBorder(range: { x?: number; y?: number; width?: number; height?: number }, edges: ReadonlySet<BorderEdges>, value?: boolean): void;
 *   mergeCells(x: number, y: number, width: number, height?: number): void;
 *   unmergeCells(x: number, y: number): void;
 *   loadCSV(text: string): void;
 * }} WorksheetEdit
 */

const NUMBER_FORMATS = ['currency', 'date', 'datetime', 'duration', 'percent'];
/**
 * `align` is `undefined` when not set: text is aligned left, numbers right and booleans center.
 * `highlight` is the name of a highlight (background/foreground color combination), `undefined` when not set.
 * `format` is how numbers are displayed, `undefined` when not set (as-is).
 * @typedef {{ bold: boolean; italic: boolean; underline: boolean; strikethrough: boolean; align: CellAlign | undefined; highlight: string | undefined; format: CellNumberFormat | undefined }} CellFormatting
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
  highlight: undefined,
  format: undefined
});

/**
 * A single worksheet of 50 columns by 10,000 rows (by default). Values are stored sparsely.
 * All coordinates (`x` for column, `y` for row) are zero-based.
 *
 * All modifications are made in a transaction, through the `edit` object passed to `edit()`:
 *
 * ```js
 * worksheet.edit(edit => {
 *   edit.setValues([['a', 'b']]);
 *   edit.setFormattings([[{ bold: true }]]);
 * });
 * ```
 *
 * When the transaction ends, a single `change` event (`WorksheetChangeEvent`) is dispatched with all changes, in order.
 * Each change has a `type` (`'value'`, `'formatting'`, `'border'`, `'merge'` or `'columnwidth'`) and carries both the
 * new and previous values, so listeners (e.g. an undo history) know exactly what was modified. If the callback throws,
 * all of its changes are rolled back and no event is dispatched.
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

  /**
   * Merged cells, keyed by `y * columnCount + x` of the top-left cell, with the number of merged columns and rows.
   * Merged cells never overlap.
   * @type {Map<number, CellMergeSize>}
   */
  #merges = new Map();

  /** @type {number[]} */
  #columnWidths;

  /** Changes of the current transaction, `undefined` outside of `edit()`. @type {{ changes: WorksheetChange[]; load: boolean } | undefined} */
  #transaction;

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
   * Sets the width of a column in pixels. Records a `columnwidth` change if changed.
   *
   * @param {number} x
   * @param {number} width
   */
  #setColumnWidth(x, width) {
    this.#assertCell(x, 0);

    if (!Number.isFinite(width) || width < 0) {
      throw new RangeError('Width must be a non-negative finite number.');
    } else if (this.#columnWidths[x] === width) {
      return;
    }

    const previousWidth = this.#columnWidths[x];

    this.#columnWidths[x] = width;
    this.#record({ type: 'columnwidth', x, width, previousWidth });
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
  #setValue(x, y, valueOrUpdater) {
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

    this.#record({ type: 'value', x, y, value, previousValue });
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
  #setValues(valuesOrUpdater, x = 0, y = 0, width, height) {
    this.#assertCell(x, y);

    if (typeof valuesOrUpdater === 'function') {
      if (!Number.isInteger(width) || !Number.isInteger(height) || width < 0 || height < 0) {
        throw new RangeError('Width and height must be non-negative integers.');
      }

      const right = Math.min(x + width, this.columnCount);
      const bottom = Math.min(y + height, this.rowCount);

      for (let cellY = y; cellY < bottom; cellY++) {
        for (let cellX = x; cellX < right; cellX++) {
          this.#setValue(cellX, cellY, valuesOrUpdater);
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
        this.#setValue(x + columnOffset, y + rowOffset, rowValues[columnOffset]);
      }
    }
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

    this.#record({ type: 'border', x, y, edge: side, border: value, previousBorder: previous[side] });
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
   * Records a `border` change (with `x` and `y` of the cell storing the edge) for each stored edge changed.
   *
   * `edges` is a `Set` of edges to apply together, for example `new Set(['top', 'left'])`. `'outside'` is a shortcut
   * of `'left'`, `'right'`, `'top'` and `'bottom'`.
   *
   * @param {{ x?: number; y?: number; width?: number; height?: number }} range
   * @param {ReadonlySet<BorderEdges>} edges
   * @param {boolean} [value] `false` to unset the borders.
   */
  #setBorder(range, edges, value = true) {
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
          // Merged cells have no inner borders.
          x > left && !(value && this.#isInsideMerge(x, y, 'left')) && this.#setStoredEdge(x, y, 'left', value);
          y > top && !(value && this.#isInsideMerge(x, y, 'top')) && this.#setStoredEdge(x, y, 'top', value);
        }

        set.has('left') &&
          x === left &&
          !(value && this.#isInsideMerge(x, y, 'left')) &&
          this.#setStoredEdge(x, y, 'left', value);
        set.has('top') && y === top && !(value && this.#isInsideMerge(x, y, 'top')) && this.#setStoredEdge(x, y, 'top', value);
        set.has('right') &&
          x === right &&
          x + 1 < this.columnCount &&
          !(value && this.#isInsideMerge(x + 1, y, 'left')) &&
          this.#setStoredEdge(x + 1, y, 'left', value);
        set.has('bottom') &&
          y === bottom &&
          y + 1 < this.rowCount &&
          !(value && this.#isInsideMerge(x, y + 1, 'top')) &&
          this.#setStoredEdge(x, y + 1, 'top', value);
      }
    }
  }

  /**
   * Sets the formatting of a cell. Missing properties are set to default. Setting `null`, `undefined` or the default
   * formatting removes the formatting. Empty cells cannot be formatted and are left as-is. Records a
   * `formatting` change if changed.
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

    const format = input?.format || undefined;

    if (format !== undefined && !NUMBER_FORMATS.includes(format)) {
      throw new TypeError(`Format must be one of ${NUMBER_FORMATS.map(name => `"${name}"`).join(', ')}.`);
    }

    /** @type {CellFormatting} */
    const formatting = {
      bold: !!input?.bold,
      italic: !!input?.italic,
      underline: !!input?.underline,
      strikethrough: !!input?.strikethrough,
      align,
      highlight,
      format
    };

    const previous = this.#formattings.get(key) || DEFAULT_FORMATTING;
    const isEqual = (/** @type {CellFormatting} */ a, /** @type {CellFormatting} */ b) =>
      a.bold === b.bold &&
      a.italic === b.italic &&
      a.underline === b.underline &&
      a.strikethrough === b.strikethrough &&
      a.align === b.align &&
      a.highlight === b.highlight &&
      a.format === b.format;

    if (isEqual(previous, formatting)) {
      return;
    }

    isEqual(formatting, DEFAULT_FORMATTING) ? this.#formattings.delete(key) : this.#formattings.set(key, formatting);

    this.#record({ type: 'formatting', x, y, formatting: { ...formatting }, previousFormatting: { ...previous } });
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
  #setFormattings(formattingsOrUpdater, x = 0, y = 0, width, height) {
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
   * Returns merged cells intersecting a range (the whole worksheet by default), in no particular order. Cost is
   * proportional to the number of merged cells, not to the size of the worksheet.
   *
   * @param {number} [x]
   * @param {number} [y]
   * @param {number} [width]
   * @param {number} [height]
   * @returns {CellMerge[]}
   */
  getMerges(x = 0, y = 0, width = this.columnCount - x, height = this.rowCount - y) {
    /** @type {CellMerge[]} */
    const merges = [];

    for (const [key, size] of this.#merges) {
      const left = key % this.columnCount;
      const top = Math.floor(key / this.columnCount);

      left < x + width &&
        left + size.width > x &&
        top < y + height &&
        top + size.height > y &&
        merges.push({ x: left, y: top, ...size });
    }

    return merges;
  }

  /**
   * Returns `true` if the edge of a cell is inside a merged cell.
   *
   * @param {number} x
   * @param {number} y
   * @param {'left' | 'top'} side
   */
  #isInsideMerge(x, y, side) {
    const merge = this.getMerges(x, y, 1, 1)[0];

    return !!merge && (side === 'left' ? merge.x !== x : merge.y !== y);
  }

  /**
   * @param {number} x
   * @param {number} y
   * @param {CellMergeSize | undefined} size `undefined` to unmerge.
   */
  #setMerge(x, y, size) {
    const key = y * this.columnCount + x;
    const previousMerge = this.#merges.get(key);

    if (previousMerge?.width === size?.width && previousMerge?.height === size?.height) {
      return;
    }

    if (size) {
      this.#assertCell(x + size.width - 1, y + size.height - 1);

      const overlap = this.getMerges(x, y, size.width, size.height).find(merge => merge.x !== x || merge.y !== y);

      if (overlap) {
        throw new Error(`Merged cells must not overlap, the merged cell at (${overlap.x}, ${overlap.y}) overlaps.`);
      }

      this.#merges.set(key, { ...size });
    } else {
      this.#merges.delete(key);
    }

    this.#record({
      type: 'merge',
      x,
      y,
      merge: size && { ...size },
      previousMerge: previousMerge && { ...previousMerge }
    });
  }

  /**
   * Merges a range of cells into a single cell. Merged cells overlapping the range are merged together.
   *
   * Non-empty values are concatenated (row by row, separated by a space) into the top-left cell and other cells are
   * cleared. The formatting of the first non-empty cell is used. Borders between the merged cells are removed.
   *
   * @param {number} x
   * @param {number} y
   * @param {number} width
   * @param {number} [height]
   */
  #mergeCells(x, y, width, height = 1) {
    if (!Number.isInteger(width) || width < 1) {
      throw new RangeError('Width must be a positive integer.');
    } else if (!Number.isInteger(height) || height < 1) {
      throw new RangeError('Height must be a positive integer.');
    }

    this.#assertCell(x, y);
    this.#assertCell(x + width - 1, y + height - 1);

    let left = x;
    let top = y;
    let right = x + width - 1;
    let bottom = y + height - 1;

    // Expanding to cover a merged cell may partially cover more merged cells, so it repeats until nothing changes.
    for (let changed = true; changed; ) {
      changed = false;

      for (const merge of this.getMerges(left, top, right - left + 1, bottom - top + 1)) {
        if (merge.x < left || merge.y < top || merge.x + merge.width - 1 > right || merge.y + merge.height - 1 > bottom) {
          left = Math.min(left, merge.x);
          top = Math.min(top, merge.y);
          right = Math.max(right, merge.x + merge.width - 1);
          bottom = Math.max(bottom, merge.y + merge.height - 1);
          changed = true;
        }
      }
    }

    if (left === right && top === bottom) {
      return;
    }

    /** @type {CellValue[]} */
    const values = [];
    /** @type {CellFormatting | undefined} */
    let formatting;

    for (const merge of this.getMerges(left, top, right - left + 1, bottom - top + 1)) {
      this.#setMerge(merge.x, merge.y, undefined);
    }

    for (let row = top; row <= bottom; row++) {
      for (let column = left; column <= right; column++) {
        const value = this.#rows[row]?.[column];

        if (value !== undefined && value !== '') {
          values.push(value);
          formatting ||= this.getFormattings(column, row, 1, 1)[0][0];
        }
      }
    }

    this.#setValue(left, top, values.length > 1 ? values.map(formatValue).join(' ') : values[0] ?? '');
    formatting && this.#setFormatting(left, top, formatting);

    for (let row = top; row <= bottom; row++) {
      for (let column = left; column <= right; column++) {
        (column !== left || row !== top) && this.#setValue(column, row, '');
        column > left && this.#setStoredEdge(column, row, 'left', false);
        row > top && this.#setStoredEdge(column, row, 'top', false);
      }
    }

    this.#setMerge(left, top, { width: right - left + 1, height: bottom - top + 1 });
  }

  /**
   * Unmerges the merged cell containing a cell. The content stays in the top-left cell.
   *
   * @param {number} x
   * @param {number} y
   */
  #unmergeCells(x, y) {
    this.#assertCell(x, y);

    const merge = this.getMerges(x, y, 1, 1)[0];

    merge && this.#setMerge(merge.x, merge.y, undefined);
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

    worksheet.edit(edit => edit.loadCSV(text));

    return worksheet;
  }

  /**
   * Loads values from CSV into this worksheet. The `change` event of the transaction has `load` set to `true`, as
   * loading is not meant to be undoable.
   *
   * @param {string} text
   */
  #loadCSV(text) {
    this.#setValues(parseCSV(text));
    /** @type {{ load: boolean }} */ (this.#transaction).load = true;
  }

  /** @param {WorksheetChange} change */
  #record(change) {
    /** @type {{ changes: WorksheetChange[] }} */ (this.#transaction).changes.push(change);
  }

  /**
   * Modifies the worksheet in a transaction. The callback receives an `edit` object with all editing functions, which
   * is only usable until the callback returns. When the transaction ends, a single `change` event
   * (`WorksheetChangeEvent`) is dispatched with all changes, unless nothing changed.
   *
   * If the callback throws, all of its changes are rolled back, no event is dispatched, and the error is rethrown.
   * Calling `edit()` inside a transaction joins it.
   *
   * @template T
   * @param {(edit: WorksheetEdit) => T} callback
   * @returns {T} The return value of the callback.
   */
  edit(callback) {
    let active = true;

    /** @type {Record<keyof WorksheetEdit, Function>} */
    const functions = {
      setColumnWidth: this.#setColumnWidth,
      setValue: this.#setValue,
      setValues: this.#setValues,
      setFormattings: this.#setFormattings,
      setBorder: this.#setBorder,
      mergeCells: this.#mergeCells,
      unmergeCells: this.#unmergeCells,
      loadCSV: this.#loadCSV
    };

    const edit = /** @type {WorksheetEdit} */ (
      Object.freeze(
        Object.fromEntries(
          Object.entries(functions).map(([name, fn]) => [
            name,
            (/** @type {any[]} */ ...args) => {
              if (!active) {
                throw new Error('The edit object can only be used inside the edit() callback.');
              }

              fn.apply(this, args);
            }
          ])
        )
      )
    );

    if (this.#transaction) {
      try {
        return callback(edit);
      } finally {
        active = false;
      }
    }

    const transaction = (this.#transaction = { changes: [], load: false });
    /** @type {T} */
    let result;

    try {
      result = callback(edit);
    } catch (error) {
      this.#rollback(transaction.changes);

      throw error;
    } finally {
      active = false;
      this.#transaction = undefined;
    }

    transaction.changes.length &&
      this.dispatchEvent(
        new WorksheetChangeEvent('change', { changes: transaction.changes, load: transaction.load, worksheet: this })
      );

    return result;
  }

  /**
   * Reverts changes of a failed transaction, last change first.
   *
   * @param {WorksheetChange[]} changes
   */
  #rollback(changes) {
    for (const change of [...changes].reverse()) {
      if (change.type === 'columnwidth') {
        this.#columnWidths[change.x] = change.previousWidth;
      } else if (change.type === 'border') {
        this.#setStoredEdge(change.x, change.y, change.edge, change.previousBorder);
      } else if (change.type === 'merge') {
        this.#setMerge(change.x, change.y, change.previousMerge);
      } else if (change.type === 'formatting') {
        this.#setFormatting(change.x, change.y, change.previousFormatting);
      } else {
        this.#setValue(change.x, change.y, change.previousValue);
      }
    }
  }
}
