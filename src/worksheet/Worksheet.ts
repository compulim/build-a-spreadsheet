import { formatValue, parseCSV, stringifyCSV } from '../csv.ts';
import { WorksheetChangeEvent } from './WorksheetChangeEvent.ts';
import { WorksheetHistory } from './WorksheetHistory.ts';
import { type CellPosition, type CellRange, WorksheetSelection } from './WorksheetSelection.ts';
import {
  boolean,
  finite,
  integer,
  is,
  looseObject,
  minValue,
  nullish,
  number,
  picklist,
  pipe,
  safeParse,
  string,
  union
} from 'valibot';
import { type Highlight, highlightSchema, parseHighlight, stringifyHighlight } from './highlight.ts';
import { COLUMN_COUNT, COLUMN_WIDTH, ROW_COUNT } from './constants.ts';
import mathRandomString from './private/mathRandomString.ts';

type CellValue = boolean | number | string;
type CellBorder = { top: boolean; right: boolean; bottom: boolean; left: boolean };
type BorderEdges = 'left' | 'right' | 'top' | 'bottom' | 'outside' | 'inside';
type CellValueChange = { type: 'value'; x: number; y: number; value: CellValue; previousValue: CellValue };
type CellFormattingChange = {
  type: 'formatting';
  x: number;
  y: number;
  formatting: CellFormatting;
  previousFormatting: CellFormatting;
};
type CellMergeSize = { width: number; height: number };
type CellMergeChange = {
  type: 'merge';
  x: number;
  y: number;
  merge: CellMergeSize | undefined;
  previousMerge: CellMergeSize | undefined;
};
type CellMerge = { x: number; y: number; width: number; height: number };
type CellBorderChange = {
  type: 'border';
  x: number;
  y: number;
  edge: 'left' | 'top';
  border: boolean;
  previousBorder: boolean;
};
type CellChange = CellValueChange | CellFormattingChange | CellMergeChange | CellBorderChange;
type WorksheetChange = CellChange;
type CellInput = CellValue | null | undefined;
type WorksheetUndo = { undo(): WorksheetRedo };
type WorksheetRedo = { redo(): WorksheetUndo };
type CellUpdater = (value: CellValue | undefined, x: number, y: number) => CellInput;
type CellAlign = 'left' | 'center' | 'right';
type CellNumberFormat = 'currency' | 'date' | 'datetime' | 'duration' | 'percent';
/**
 * Editing functions, only usable inside the `edit()` callback.
 */
type WorksheetEdit = {
  setValue(x: number, y: number, valueOrUpdater: CellInput | CellUpdater): void;
  setValues(
    valuesOrUpdater: CellInput[][] | CellUpdater,
    x?: number,
    y?: number,
    width?: number,
    height?: number
  ): void;
  setFormattings(
    formattingsOrUpdater: FormattingInput[][] | FormattingUpdater,
    x?: number,
    y?: number,
    width?: number,
    height?: number
  ): void;
  setBorder(
    range: { x?: number; y?: number; width?: number; height?: number },
    edges: ReadonlySet<BorderEdges>,
    value?: boolean
  ): void;
  mergeCells(x: number, y: number, width: number, height?: number): void;
  unmergeCells(x: number, y: number): void;
  loadCSV(text: string): void;
};

const NUMBER_FORMATS = ['currency', 'date', 'datetime', 'duration', 'percent'] as const;

const integerSchema = pipe(number(), integer());
const nonNegativeIntegerSchema = pipe(number(), integer(), minValue(0));
const nonNegativeFiniteSchema = pipe(number(), finite(), minValue(0));
const cellInputSchema = nullish(union([boolean(), pipe(number(), finite()), string()]));
const formattingInputSchema = nullish(
  looseObject({
    bold: nullish(boolean()),
    italic: nullish(boolean()),
    underline: nullish(boolean()),
    strikethrough: nullish(boolean()),
    align: nullish(picklist(['left', 'center', 'right'])),
    highlight: nullish(highlightSchema),
    format: nullish(picklist(NUMBER_FORMATS))
  })
);

/**
 * `align` is `undefined` when not set: text is aligned left, numbers right and booleans center.
 * `highlight` is the background and text colors in hex, either can be empty to keep the default (e.g.
 * `{ background: '#c6efce', color: '' }`), `undefined` when not set.
 * `format` is how numbers are displayed, `undefined` when not set (as-is).
 */
type CellFormatting = {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  align: CellAlign | undefined;
  highlight: Highlight | undefined;
  format: CellNumberFormat | undefined;
};
/** Formatting for internal storage, `highlight` is stored as `"#background/#color"`. */
type StoredFormatting = Omit<CellFormatting, 'highlight'> & { highlight: string | undefined };
type FormattingInput = Partial<CellFormatting> | null | undefined;
type FormattingUpdater = (formatting: CellFormatting, x: number, y: number) => FormattingInput;

const DEFAULT_FORMATTING: Readonly<StoredFormatting> = Object.freeze({
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  align: undefined,
  highlight: undefined,
  format: undefined
});

function toCellFormatting({ highlight, ...formatting }: StoredFormatting): CellFormatting {
  return { ...formatting, highlight: parseHighlight(highlight) };
}

function isSameSelection(
  { activeCell, range }: WorksheetSelection,
  { column, row }: CellPosition,
  { left, top, right, bottom }: CellRange
): boolean {
  return (
    activeCell.column === column &&
    activeCell.row === row &&
    range.left === left &&
    range.top === top &&
    range.right === right &&
    range.bottom === bottom
  );
}

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
 * Each change has a `type` (`'value'`, `'formatting'`, `'border'` or `'merge'`) and carries both the
 * new and previous values, so listeners (e.g. an undo history) know exactly what was modified. If the callback throws,
 * all of its changes are rolled back and no event is dispatched.
 *
 * Every transaction with changes gives the worksheet a new `version` (a GUID). `edit()` returns an object to undo
 * the transaction, only callable while the worksheet is still at the version right after it. Undoing restores the
 * previous version and returns an object to redo, only callable while the worksheet is at that previous version.
 *
 * The worksheet also keeps an undo history of its transactions (`undo()`, `redo()`, `canUndo` and `canRedo`) and the
 * `selection`. The selection is not data: it is set directly, outside of `edit()`, and does not change the version.
 *
 * Events: `change` (`WorksheetChangeEvent`), `columnwidthchange` (`CustomEvent`), `selectionchange` and
 * `historychange` (undo/redo availability may have changed, dispatched asynchronously).
 */
class Worksheet extends EventTarget {
  #rows: (CellValue[] | undefined)[] = [];

  /**
   * Left and top borders of cells, keyed by `y * columnCount + x`. The right and bottom borders of a cell are the left
   * border of the cell on the right and the top border of the cell below.
   */
  #borders: Map<number, { top: boolean; left: boolean }> = new Map();

  /**
   * Formatting of cells, keyed by `y * columnCount + x`. Unformatted cells are not stored.
   */
  #formattings: Map<number, StoredFormatting> = new Map();

  /**
   * Merged cells, keyed by `y * columnCount + x` of the top-left cell, with the number of merged columns and rows.
   * Merged cells never overlap.
   */
  #merges: Map<number, CellMergeSize> = new Map();

  #columnWidths: number[];

  /** Changes of the current transaction, `undefined` outside of `edit()`. */
  #transaction: { changes: WorksheetChange[]; load: boolean } | undefined;

  /** Changes on every transaction with changes, restored by undo and redo. */
  #version: string = mathRandomString();

  /** Undo history, built by subscribing to the `change` event. */
  #history: WorksheetHistory;

  #historyChangeScheduled = false;

  #selection: WorksheetSelection = new WorksheetSelection({ column: 0, row: 0 });

  /** Range of the last selection set, before it was expanded to cover merged cells. */
  #requestedRange: CellRange = { left: 0, top: 0, right: 0, bottom: 0 };

  readonly columnCount: number;
  readonly rowCount: number;

  constructor({
    columnCount = COLUMN_COUNT,
    rowCount = ROW_COUNT,
    columnWidth = COLUMN_WIDTH
  }: { columnCount?: number; rowCount?: number; columnWidth?: number } = {}) {
    super();

    this.columnCount = columnCount;
    this.rowCount = rowCount;

    this.#columnWidths = new Array(columnCount).fill(columnWidth);

    // The history subscribes first, so other `change` listeners see the history after it recorded the change.
    this.#history = new WorksheetHistory(this);
    this.addEventListener('change', this.#handleChange);
  }

  #handleChange = (event: Event) => {
    const { changes } = event as WorksheetChangeEvent;

    // Merged cells may have been added or removed, the selection must cover whole merged cells.
    changes.some(({ type }) => type === 'merge') &&
      this.#applySelection(this.#selection.activeCell, this.#requestedRange);

    this.#scheduleHistoryChange();
  };

  #scheduleHistoryChange() {
    if (!this.#historyChangeScheduled) {
      this.#historyChangeScheduled = true;

      queueMicrotask(() => {
        this.#historyChangeScheduled = false;
        this.dispatchEvent(new Event('historychange'));
      });
    }
  }

  /**
   * Current selection. Returns the same instance until the selection changes.
   */
  get selection(): WorksheetSelection {
    return this.#selection;
  }

  /**
   * Selects cells. The selection is clamped to the worksheet, the active cell is moved to the top-left cell of its
   * merged cell and the range is expanded to cover merged cells partially in it. The getter returns the very same
   * instance, unless the selection was adjusted.
   *
   * Dispatches `selectionchange` if the selection changed.
   */
  set selection(selection: WorksheetSelection) {
    const { left, top, right, bottom } = selection.range;
    const topLeft = this.#clamp({ column: left, row: top });
    const bottomRight = this.#clamp({ column: right, row: bottom });

    this.#requestedRange = {
      left: topLeft.column,
      top: topLeft.row,
      right: bottomRight.column,
      bottom: bottomRight.row
    };
    this.#applySelection(this.#clamp(selection.activeCell), this.#requestedRange, selection);
  }

  #applySelection(activeCell: CellPosition, requestedRange: CellRange, candidate?: WorksheetSelection) {
    const mergeStart = this.getMerges(activeCell.column, activeCell.row, 1, 1)[0];
    const nextActiveCell = mergeStart ? { column: mergeStart.x, row: mergeStart.y } : activeCell;
    const range = this.#expandRange({
      left: Math.min(requestedRange.left, nextActiveCell.column),
      top: Math.min(requestedRange.top, nextActiveCell.row),
      right: Math.max(requestedRange.right, nextActiveCell.column),
      bottom: Math.max(requestedRange.bottom, nextActiveCell.row)
    });

    if (candidate && isSameSelection(candidate, nextActiveCell, range)) {
      const changed = !isSameSelection(this.#selection, nextActiveCell, range);

      this.#selection = candidate;
      changed && this.dispatchEvent(new Event('selectionchange'));
    } else if (!isSameSelection(this.#selection, nextActiveCell, range)) {
      this.#selection = new WorksheetSelection(nextActiveCell, range);
      this.dispatchEvent(new Event('selectionchange'));
    }
  }

  #clamp({ column, row }: CellPosition): CellPosition {
    return {
      column: Math.max(0, Math.min(this.columnCount - 1, column)),
      row: Math.max(0, Math.min(this.rowCount - 1, row))
    };
  }

  /**
   * Returns the range expanded to cover all merged cells partially in it.
   */
  #expandRange({ left, top, right, bottom }: CellRange): CellRange {
    // Expanding may cover more merged cells, so it repeats until nothing changes.
    for (let changed = true; changed;) {
      changed = false;

      for (const { x, y, width, height } of this.getMerges(left, top, right - left + 1, bottom - top + 1)) {
        if (x < left || y < top || x + width - 1 > right || y + height - 1 > bottom) {
          left = Math.min(left, x);
          top = Math.min(top, y);
          right = Math.max(right, x + width - 1);
          bottom = Math.max(bottom, y + height - 1);
          changed = true;
        }
      }
    }

    return { left, top, right, bottom };
  }

  /** Whether there are changes to undo. */
  get canUndo(): boolean {
    return this.#history.canUndo;
  }

  /** Whether there are undone changes to redo. */
  get canRedo(): boolean {
    return this.#history.canRedo;
  }

  /**
   * Reverts the last transaction and selects the changed cells.
   *
   * Returns `false` if there is nothing to undo.
   */
  undo(): boolean {
    return this.#undoOrRedo('undo');
  }

  /**
   * Re-applies the last undone transaction and selects the changed cells.
   *
   * Returns `false` if there is nothing to redo.
   */
  redo(): boolean {
    return this.#undoOrRedo('redo');
  }

  #undoOrRedo(action: 'undo' | 'redo'): boolean {
    const cells: CellPosition[] = [];

    const handleChange = (event: Event) => {
      for (const change of (event as WorksheetChangeEvent).changes) {
        if (change.type === 'merge') {
          // All cells of a merged cell are involved, not only the cells with content.
          const { width, height } = (change.merge || change.previousMerge)!;

          cells.push({ column: change.x, row: change.y }, { column: change.x + width - 1, row: change.y + height - 1 });
        } else if (change.type !== 'border') {
          cells.push({ column: change.x, row: change.y });
        }
      }
    };

    // The replay is a single transaction, dispatching a single `change` event synchronously.
    this.addEventListener('change', handleChange);

    try {
      if (!this.#history[action]()) {
        return false;
      }
    } finally {
      this.removeEventListener('change', handleChange);
    }

    if (cells.length) {
      const columns = cells.map(({ column }) => column);
      const rows = cells.map(({ row }) => row);
      const left = Math.min(...columns);
      const top = Math.min(...rows);

      this.selection = new WorksheetSelection(
        { column: left, row: top },
        { left, top, right: Math.max(...columns), bottom: Math.max(...rows) }
      );
    }

    this.#scheduleHistoryChange();

    return true;
  }

  /** Version of the worksheet, a GUID replaced on every transaction with changes. */
  get version(): string {
    return this.#version;
  }

  #assertCell(x: number, y: number): void {
    if (!is(integerSchema, x) || x < 0 || x >= this.columnCount) {
      throw new RangeError(`X must be an integer between 0 and ${this.columnCount - 1}.`);
    } else if (!is(integerSchema, y) || y < 0 || y >= this.rowCount) {
      throw new RangeError(`Y must be an integer between 0 and ${this.rowCount - 1}.`);
    }
  }

  /**
   * Returns the width of a column in pixels.
   */
  getColumnWidth(x: number): number {
    this.#assertCell(x, 0);

    return this.#columnWidths[x]!;
  }

  /**
   * Sets the width of a column in pixels. Not part of `edit()`: it is not undoable and does not change the `version`.
   * Dispatches a `columnwidthchange` event (`CustomEvent` with `detail` of `{ x, width, previousWidth }`) if changed.
   */
  setColumnWidth(x: number, width: number): void {
    this.#assertCell(x, 0);

    if (!is(nonNegativeFiniteSchema, width)) {
      throw new RangeError('Width must be a non-negative finite number.');
    } else if (this.#columnWidths[x] === width) {
      return;
    }

    const previousWidth = this.#columnWidths[x];

    this.#columnWidths[x] = width;
    this.dispatchEvent(new CustomEvent('columnwidthchange', { detail: { x, width, previousWidth } }));
  }

  /**
   * Returns the value of a cell, empty cells return `undefined`.
   */
  getValue(x: number, y: number): CellValue | undefined {
    this.#assertCell(x, y);

    return this.#rows[y]?.[x];
  }

  /**
   * Sets the value of a cell. Strings are trimmed. Setting an empty or whitespace-only string, `null` or `undefined` clears the cell.
   *
   * @param valueOrUpdater A value, or a function `(value, x, y) => nextValue`.
   */
  #setValue(x: number, y: number, valueOrUpdater: CellInput | CellUpdater): void {
    this.#assertCell(x, y);

    let value = typeof valueOrUpdater === 'function' ? valueOrUpdater(this.#rows[y]?.[x], x, y) : valueOrUpdater;

    if (!is(cellInputSchema, value)) {
      throw new TypeError(
        typeof value === 'number' ? 'Number value must be finite.' : 'Value must be a boolean, number or string.'
      );
    }

    if (value === null || value === undefined) {
      value = '';
    } else if (typeof value === 'string') {
      value = value.trim();
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
  #getUsedSize(): { width: number; height: number } {
    let height = 0;
    let width = 0;

    this.#rows.forEach((cells, y) => {
      height = y + 1;
      cells!.forEach((_, x) => {
        width = Math.max(width, x + 1);
      });
    });

    return { width, height };
  }

  /**
   * Returns values of a range as a 2D array (`values[y][x]`). Empty cells are returned as `''`.
   */
  getValues(x: number, y: number, width: number, height: number): CellValue[][] {
    if (!is(nonNegativeIntegerSchema, width) || !is(nonNegativeIntegerSchema, height)) {
      throw new RangeError('Width and height must be non-negative integers.');
    }

    if (!width || !height) {
      return [];
    }

    this.#assertCell(x, y);
    this.#assertCell(x + width - 1, y + height - 1);

    const values: CellValue[][] = [];

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
   * @param width Required with an updater function.
   * @param height Required with an updater function.
   */
  #setValues(
    valuesOrUpdater: CellInput[][] | CellUpdater,
    x: number = 0,
    y: number = 0,
    width?: number,
    height?: number
  ): void {
    this.#assertCell(x, y);

    if (typeof valuesOrUpdater === 'function') {
      if (!is(nonNegativeIntegerSchema, width) || !is(nonNegativeIntegerSchema, height)) {
        throw new RangeError('Width and height must be non-negative integers.');
      }

      const right = Math.min(x + width!, this.columnCount);
      const bottom = Math.min(y + height!, this.rowCount);

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
   */
  getBorder(x: number, y: number): CellBorder {
    this.#assertCell(x, y);

    const own = this.#borders.get(y * this.columnCount + x);

    return {
      top: !!own?.top,
      right: x + 1 < this.columnCount && !!this.#borders.get(y * this.columnCount + x + 1)?.left,
      bottom: y + 1 < this.rowCount && !!this.#borders.get((y + 1) * this.columnCount + x)?.top,
      left: !!own?.left
    };
  }

  #setStoredEdge(x: number, y: number, side: 'left' | 'top', value: boolean): void {
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
   * @param value `false` to unset the borders.
   */
  #setBorder(
    range: { x?: number; y?: number; width?: number; height?: number },
    edges: ReadonlySet<BorderEdges>,
    value: boolean = true
  ): void {
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
        set.has('top') &&
          y === top &&
          !(value && this.#isInsideMerge(x, y, 'top')) &&
          this.#setStoredEdge(x, y, 'top', value);
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
   * @param formattingOrUpdater A formatting, or a function `(formatting, x, y) => nextFormatting`.
   */
  #setFormatting(x: number, y: number, formattingOrUpdater: FormattingInput | FormattingUpdater): void {
    if ((this.#rows[y]?.[x] ?? '') === '') {
      return;
    }

    const key = y * this.columnCount + x;
    const input =
      typeof formattingOrUpdater === 'function'
        ? formattingOrUpdater(toCellFormatting(this.#formattings.get(key) || DEFAULT_FORMATTING), x, y)
        : formattingOrUpdater;

    const result = safeParse(formattingInputSchema, input);

    if (!result.success) {
      const path = result.issues[0].path?.map(({ key }) => String(key)).join('.');

      throw new TypeError(`Invalid formatting${path ? ` at "${path}"` : ''}: ${result.issues[0].message}`);
    }

    const { output } = result;
    const formatting: StoredFormatting = {
      bold: !!output?.bold,
      italic: !!output?.italic,
      underline: !!output?.underline,
      strikethrough: !!output?.strikethrough,
      align: output?.align ?? undefined,
      highlight: output?.highlight ? stringifyHighlight(output.highlight) : undefined,
      format: output?.format ?? undefined
    };

    const previous = this.#formattings.get(key) || DEFAULT_FORMATTING;
    const isEqual = (a: StoredFormatting, b: StoredFormatting): boolean =>
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

    this.#record({
      type: 'formatting',
      x,
      y,
      formatting: toCellFormatting(formatting),
      previousFormatting: toCellFormatting(previous)
    });
  }

  /**
   * Returns formattings of a range as a 2D array (`formattings[y][x]`). Unformatted cells return the default
   * formatting (no styles, `align` of `undefined`).
   */
  getFormattings(x: number, y: number, width: number, height: number): CellFormatting[][] {
    if (!is(nonNegativeIntegerSchema, width) || !is(nonNegativeIntegerSchema, height)) {
      throw new RangeError('Width and height must be non-negative integers.');
    }

    if (!width || !height) {
      return [];
    }

    this.#assertCell(x, y);
    this.#assertCell(x + width - 1, y + height - 1);

    const formattings: CellFormatting[][] = [];

    for (let rowOffset = 0; rowOffset < height; rowOffset++) {
      const row = [];

      for (let columnOffset = 0; columnOffset < width; columnOffset++) {
        row.push(
          toCellFormatting(
            this.#formattings.get((y + rowOffset) * this.columnCount + x + columnOffset) || DEFAULT_FORMATTING
          )
        );
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
   * @param width Required with an updater function.
   * @param height Required with an updater function.
   */
  #setFormattings(
    formattingsOrUpdater: FormattingInput[][] | FormattingUpdater,
    x: number = 0,
    y: number = 0,
    width?: number,
    height?: number
  ): void {
    this.#assertCell(x, y);

    if (typeof formattingsOrUpdater === 'function') {
      if (!is(nonNegativeIntegerSchema, width) || !is(nonNegativeIntegerSchema, height)) {
        throw new RangeError('Width and height must be non-negative integers.');
      }

      const right = Math.min(x + width!, this.columnCount);
      const bottom = Math.min(y + height!, this.rowCount);

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
   */
  getMerges(
    x: number = 0,
    y: number = 0,
    width: number = this.columnCount - x,
    height: number = this.rowCount - y
  ): CellMerge[] {
    const merges: CellMerge[] = [];

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
   */
  #isInsideMerge(x: number, y: number, side: 'left' | 'top'): boolean {
    const merge = this.getMerges(x, y, 1, 1)[0];

    return !!merge && (side === 'left' ? merge.x !== x : merge.y !== y);
  }

  /**
   * @param size `undefined` to unmerge.
   */
  #setMerge(x: number, y: number, size: CellMergeSize | undefined): void {
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
   */
  #mergeCells(x: number, y: number, width: number, height: number = 1): void {
    if (!is(nonNegativeIntegerSchema, width) || width < 1) {
      throw new RangeError('Width must be a positive integer.');
    } else if (!is(nonNegativeIntegerSchema, height) || height < 1) {
      throw new RangeError('Height must be a positive integer.');
    }

    this.#assertCell(x, y);
    this.#assertCell(x + width - 1, y + height - 1);

    let left = x;
    let top = y;
    let right = x + width - 1;
    let bottom = y + height - 1;

    // Expanding to cover a merged cell may partially cover more merged cells, so it repeats until nothing changes.
    for (let changed = true; changed;) {
      changed = false;

      for (const merge of this.getMerges(left, top, right - left + 1, bottom - top + 1)) {
        if (
          merge.x < left ||
          merge.y < top ||
          merge.x + merge.width - 1 > right ||
          merge.y + merge.height - 1 > bottom
        ) {
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

    const values: CellValue[] = [];
    let formatting: CellFormatting | undefined;

    for (const merge of this.getMerges(left, top, right - left + 1, bottom - top + 1)) {
      this.#setMerge(merge.x, merge.y, undefined);
    }

    for (let row = top; row <= bottom; row++) {
      for (let column = left; column <= right; column++) {
        const value = this.#rows[row]?.[column];

        if (value !== undefined && value !== '') {
          values.push(value);
          formatting ||= this.getFormattings(column, row, 1, 1)[0]![0];
        }
      }
    }

    this.#setValue(left, top, values.length > 1 ? values.map(formatValue).join(' ') : (values[0] ?? ''));
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
   */
  #unmergeCells(x: number, y: number): void {
    this.#assertCell(x, y);

    const merge = this.getMerges(x, y, 1, 1)[0];

    merge && this.#setMerge(merge.x, merge.y, undefined);
  }

  /**
   * Serializes the worksheet into CSV.
   */
  toCSV(): string {
    const { width, height } = this.#getUsedSize();

    return stringifyCSV(this.getValues(0, 0, width, height));
  }

  /**
   * Creates a worksheet from CSV.
   */
  static fromCSV(text: string): Worksheet {
    const worksheet = new Worksheet();

    worksheet.edit(edit => edit.loadCSV(text));

    return worksheet;
  }

  /**
   * Loads values from CSV into this worksheet. The `change` event of the transaction has `load` set to `true`, as
   * loading is not meant to be undoable.
   */
  #loadCSV(text: string): void {
    this.#setValues(parseCSV(text));
    (this.#transaction as { load: boolean }).load = true;
  }

  #record(change: WorksheetChange): void {
    (this.#transaction as { changes: WorksheetChange[] }).changes.push(change);
  }

  /**
   * Modifies the worksheet in a transaction. The callback receives an `edit` object with all editing functions, which
   * is only usable until the callback returns. When the transaction ends, a single `change` event
   * (`WorksheetChangeEvent`) is dispatched with all changes, unless nothing changed.
   *
   * If the callback throws, all of its changes are rolled back, no event is dispatched, and the error is rethrown.
   * Calling `edit()` inside a transaction joins it, the returned object cannot undo it separately.
   *
   * The returned `undo()` reverts the changes of the transaction and restores the
   * version before it. It throws if the worksheet is not at the version right after the transaction.
   */
  edit(callback: (edit: WorksheetEdit) => void): WorksheetUndo {
    let active = true;

    const functions: Record<keyof WorksheetEdit, Function> = {
      setValue: this.#setValue,
      setValues: this.#setValues,
      setFormattings: this.#setFormattings,
      setBorder: this.#setBorder,
      mergeCells: this.#mergeCells,
      unmergeCells: this.#unmergeCells,
      loadCSV: this.#loadCSV
    };

    const edit = Object.freeze(
      Object.fromEntries(
        Object.entries(functions).map(([name, fn]) => [
          name,
          (...args: any[]) => {
            if (!active) {
              throw new Error('The edit object can only be used inside the edit() callback.');
            }

            fn.apply(this, args);
          }
        ])
      )
    ) as WorksheetEdit;

    if (this.#transaction) {
      try {
        callback(edit);
      } finally {
        active = false;
      }

      return {
        undo() {
          throw new Error('A nested edit() cannot be undone separately.');
        }
      };
    }

    const transaction = (this.#transaction = { changes: [], load: false });

    try {
      callback(edit);
    } catch (error) {
      this.#rollback(transaction.changes);

      throw error;
    } finally {
      active = false;
      this.#transaction = undefined;
    }

    const { changes } = transaction;
    const previousVersion = this.#version;
    const version = changes.length ? mathRandomString() : previousVersion;

    this.#version = version;

    const undo = this.#createUndo(changes, previousVersion, version);

    changes.length &&
      this.dispatchEvent(
        new WorksheetChangeEvent('change', {
          changes,
          load: transaction.load,
          undo,
          worksheet: this
        })
      );

    return undo;
  }

  #createUndo(changes: CellChange[], previousVersion: string, version: string): WorksheetUndo {
    return {
      undo: () => {
        this.#replay(changes, true, version, previousVersion);

        return {
          redo: () => {
            this.#replay(changes, false, previousVersion, version);

            return this.#createUndo(changes, previousVersion, version);
          }
        };
      }
    };
  }

  /**
   * Reverts or re-applies cell changes, then dispatches a `change` event (without `undo`) with the changes in the order
   * they were applied.
   *
   * @param fromVersion Version the worksheet must be at.
   * @param toVersion Version of the worksheet afterward.
   */
  #replay(changes: CellChange[], isUndo: boolean, fromVersion: string, toVersion: string): void {
    if (this.#transaction) {
      throw new Error('Cannot undo or redo inside edit().');
    } else if (this.#version !== fromVersion) {
      throw new Error(`Cannot ${isUndo ? 'undo' : 'redo'}, the worksheet has changed since.`);
    }

    const ordered = isUndo ? [...changes].reverse() : changes;

    // Low-level setters record into a scratch transaction, which is discarded.
    this.#transaction = { changes: [], load: false };

    try {
      for (const change of ordered) {
        if (change.type === 'border') {
          this.#setStoredEdge(change.x, change.y, change.edge, isUndo ? change.previousBorder : change.border);
        } else if (change.type === 'merge') {
          this.#setMerge(change.x, change.y, isUndo ? change.previousMerge : change.merge);
        } else if (change.type === 'formatting') {
          this.#setFormatting(change.x, change.y, isUndo ? change.previousFormatting : change.formatting);
        } else {
          this.#setValue(change.x, change.y, isUndo ? change.previousValue : change.value);
        }
      }
    } finally {
      this.#transaction = undefined;
    }

    this.#version = toVersion;

    ordered.length &&
      this.dispatchEvent(new WorksheetChangeEvent('change', { changes: ordered, load: false, worksheet: this }));
  }

  /**
   * Reverts changes of a failed transaction, last change first.
   */
  #rollback(changes: WorksheetChange[]): void {
    for (const change of [...changes].reverse()) {
      if (change.type === 'border') {
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

export {
  type BorderEdges,
  type CellAlign,
  type CellBorder,
  type CellBorderChange,
  type CellChange,
  type CellFormatting,
  type CellFormattingChange,
  type CellInput,
  type CellMerge,
  type CellMergeChange,
  type CellMergeSize,
  type CellNumberFormat,
  type CellUpdater,
  type CellValue,
  type CellValueChange,
  type FormattingInput,
  type FormattingUpdater,
  Worksheet,
  type WorksheetChange,
  type WorksheetEdit,
  type WorksheetRedo,
  type WorksheetUndo
};
