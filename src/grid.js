import { calculate } from './calculator.js';
import { formatValue, inferValue, parseTSV, stringifyTSV } from './csv.js';
import { getHighlight } from './highlights.js';
import { formatDisplayValue } from './numberFormats.js';
import { getColumnName } from './workbook/getColumnName.js';

/**
 * @typedef {import('./workbook/Worksheet.js').Worksheet} Worksheet
 * @typedef {import('./workbook/Worksheet.js').CellValue} CellValue
 * @typedef {{ column: number; row: number }} CellPosition
 * @typedef {{ left: number; top: number; right: number; bottom: number }} CellRange
 * @typedef {{ anchor: CellPosition; focus: CellPosition; scrollLeft: number; scrollTop: number }} ViewState
 */

/**
 * Minimum number of extra rows rendered above and below the visible area. Otherwise, about one viewport of extra rows is
 * rendered on each side, so fast or momentum scrolling (which may be composited ahead of the main thread) does not show
 * blank rows.
 */
const MIN_OVERSCAN = 4;

const prefersReducedMotion =
  typeof matchMedia === 'function' ? matchMedia('(prefers-reduced-motion: reduce)') : { matches: false };

/**
 * @param {string} tagName
 * @param {string} className
 * @returns {HTMLElement}
 */
const MIN_COLUMN_WIDTH = 20;

/** Distance in pixels from a column header border where resizing can start. */
const RESIZE_HANDLE_SIZE = 8;

/** Width in pixels of cell borders set by the user. */
const BORDER_WIDTH = 3;

const BORDER_COLOR = '#000';

/** Width in pixels of the part of a border drawn past the grid line, inside the cell on the right or below. */
const BORDER_OUTSIDE_WIDTH = Math.floor((BORDER_WIDTH - 1) / 2);

/**
 * Returns the border of a cell, or `undefined` if the cell is outside of the worksheet.
 *
 * @param {Worksheet} worksheet
 * @param {number} column
 * @param {number} row
 */
function getBorderOrNone(worksheet, column, row) {
  return column >= 0 && row >= 0 && column < worksheet.columnCount && row < worksheet.rowCount
    ? worksheet.getBorder(column, row)
    : undefined;
}

/**
 * Returns `true` if there is a border on the grid line on the right of the cell (column may be -1).
 *
 * @param {Worksheet} worksheet
 * @param {number} column
 * @param {number} row
 */
function hasVerticalEdge(worksheet, column, row) {
  return !!(getBorderOrNone(worksheet, column, row)?.right || getBorderOrNone(worksheet, column + 1, row)?.left);
}

/**
 * Returns `true` if there is a border on the grid line below the cell (row may be -1).
 *
 * @param {Worksheet} worksheet
 * @param {number} column
 * @param {number} row
 */
function hasHorizontalEdge(worksheet, column, row) {
  return !!(getBorderOrNone(worksheet, column, row)?.bottom || getBorderOrNone(worksheet, column, row + 1)?.top);
}

/**
 * Returns `true` if any border ends at or passes through the bottom-right corner of the cell.
 *
 * @param {Worksheet} worksheet
 * @param {number} column
 * @param {number} row
 */
function hasJunction(worksheet, column, row) {
  return (
    hasVerticalEdge(worksheet, column, row) ||
    hasVerticalEdge(worksheet, column, row + 1) ||
    hasHorizontalEdge(worksheet, column, row) ||
    hasHorizontalEdge(worksheet, column + 1, row)
  );
}

function createElement(tagName, className) {
  const element = document.createElement(tagName);

  element.className = className;

  return element;
}

/**
 * Returns formattings embedded by this app in clipboard HTML, or `undefined` if the HTML is not from this app.
 *
 * @param {string} html
 * @returns {Partial<import('./workbook/Worksheet.js').CellFormatting>[][] | undefined}
 */
function parseClipboardFormattings(html) {
  if (!html) {
    return;
  }

  try {
    const json = new DOMParser().parseFromString(html, 'text/html').querySelector('table[data-formattings]')?.getAttribute('data-formattings');
    const formattings = json && JSON.parse(json);

    return Array.isArray(formattings) && formattings.every(Array.isArray) ? formattings : undefined;
  } catch {
    return;
  }
}

/**
 * @param {string} text
 * @returns {string}
 */
function escapeHTML(text) {
  return text.replace(
    /[&<>"']/g,
    char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char] || char
  );
}

/**
 * Virtualized spreadsheet grid. Only the visible rows (plus a few extra) are in the DOM, while the scrolling itself is
 * handled by the browser using a native scrollbar on a container sized to the full worksheet.
 */
export class SpreadsheetGrid extends EventTarget {
  /** @type {Worksheet | undefined} */
  #worksheet;

  /** @type {WeakMap<Worksheet, ViewState>} */
  #viewStates = new WeakMap();

  /** Active cell, where editing happens. @type {CellPosition} */
  #anchor = { column: 0, row: 0 };

  /** The other end of the selected range. @type {CellPosition} */
  #focus = { column: 0, row: 0 };

  #locked = false;

  /** Number of leading columns that are read-only and frozen while locked. */
  #lockedColumnCount;

  /** Number of leading rows that are read-only and frozen while locked. */
  #lockedRowCount;

  /** Copies of the locked rows, shown in a sticky pane while locked. @type {{ header: HTMLElement; cells: HTMLElement[]; row: number }[]} */
  #frozenRows = [];

  /** @type {{ element: HTMLElement; header: HTMLElement; cells: HTMLElement[]; row: number }[]} */
  #pool = [];

  /** @type {HTMLElement[]} */
  #columnHeaders = [];

  /** Whether the editor is open, and whether arrow keys should commit ("enter" mode) or move the caret ("edit" mode). */
  #editMode = /** @type {'' | 'enter' | 'edit'} */ ('');

  #renderScheduled = false;

  /** Where an in-progress smooth scroll will end. @type {{ left: number; top: number } | undefined} */
  #scrollTarget;

  /** @type {ReturnType<typeof setTimeout> | undefined} */
  #scrollEndTimeout;
  #firstRow = 0;
  #contentDirty = false;

  /** @type {'cell' | 'row' | 'column' | 'resize' | ''} */
  #dragging = '';

  /** Columns being resized and where the drag started. @type {{ columns: number[]; startX: number; startWidth: number } | undefined} */
  #resizing;

  /**
   * @param {HTMLElement} container
   * @param {{
   *   columnWidth?: number;
   *   rowHeight?: number;
   *   rowHeaderWidth?: number;
   *   columnCount?: number;
   *   rowCount?: number;
   *   lockedColumnCount?: number;
   *   lockedRowCount?: number;
   * }} [options]
   */
  constructor(
    container,
    {
      columnWidth = 100,
      rowHeight = 24,
      rowHeaderWidth = 56,
      columnCount = 50,
      rowCount = 10_000,
      lockedColumnCount = 0,
      lockedRowCount = 0
    } = {}
  ) {
    super();

    this.columnWidth = columnWidth;
    this.rowHeight = rowHeight;
    this.headerHeight = rowHeight;
    this.rowHeaderWidth = rowHeaderWidth;
    this.columnCount = columnCount;
    this.rowCount = rowCount;
    this.#lockedColumnCount = Math.max(0, Math.min(columnCount, lockedColumnCount));
    this.#lockedRowCount = Math.max(0, Math.min(rowCount, lockedRowCount));

    const scroller = (this.element = createElement('div', 'grid'));

    scroller.tabIndex = 0;
    scroller.setAttribute('role', 'grid');
    scroller.setAttribute('aria-rowcount', String(rowCount + 1));
    scroller.setAttribute('aria-colcount', String(columnCount + 1));
    scroller.style.setProperty('--column-width', `${columnWidth}px`);
    scroller.style.setProperty('--row-height', `${rowHeight}px`);
    scroller.style.setProperty('--row-header-width', `${rowHeaderWidth}px`);
    scroller.style.setProperty('--locked-row-count', String(this.#lockedRowCount));

    const content = createElement('div', 'grid__content');

    this.#content = content;

    const header = createElement('div', 'grid__header');

    header.setAttribute('role', 'row');
    header.append(createElement('div', 'grid__corner'));

    for (let column = 0; column < columnCount; column++) {
      const columnHeader = createElement('div', 'grid__column-header');

      columnHeader.textContent = getColumnName(column);
      columnHeader.dataset.column = String(column);
      columnHeader.style.width = `${columnWidth}px`;
      columnHeader.setAttribute('role', 'columnheader');
      this.#initLockableColumn(columnHeader, column);
      this.#columnHeaders.push(columnHeader);
      header.append(columnHeader);
    }

    // Locked rows are frozen by showing copies of them in a sticky pane below the header, while the body is shifted up
    // underneath it so all other coordinates stay the same.
    const frozen = createElement('div', 'grid__frozen');

    for (let row = 0; row < this.#lockedRowCount; row++) {
      const slot = this.#createRow();

      slot.element.classList.add('grid__row--lockable');
      slot.element.setAttribute('aria-hidden', 'true');
      slot.header.textContent = String(row + 1);
      slot.row = row;
      frozen.append(slot.element);
      this.#frozenRows.push(slot);
    }

    const body = (this.#body = createElement('div', 'grid__body'));

    body.style.height = `${rowCount * rowHeight}px`;

    this.#rangeOverlay = createElement('div', 'grid__range');
    this.#activeOverlay = createElement('div', 'grid__active');

    const editor = (this.#editor = /** @type {HTMLInputElement} */ (createElement('input', 'grid__editor')));

    editor.type = 'text';
    editor.spellcheck = false;
    editor.autocomplete = 'off';
    editor.hidden = true;

    body.append(this.#rangeOverlay, this.#activeOverlay, editor);
    content.append(header, frozen, body);
    scroller.append(content);
    container.append(scroller);

    // Render synchronously in the scroll event (fired once per frame), so new rows are ready before the frame is painted.
    scroller.addEventListener('scroll', () => this.#handleScroll(), { passive: true });
    scroller.addEventListener('scrollend', () => this.#clearScrollTarget(), { passive: true });

    // User-initiated scrolling interrupts any smooth scroll started by keyboard navigation.
    for (const type of ['wheel', 'touchstart']) {
      scroller.addEventListener(type, () => this.#clearScrollTarget(), { passive: true });
    }
    scroller.addEventListener('keydown', event => this.#handleKeyDown(event));
    scroller.addEventListener('pointerdown', event => this.#handlePointerDown(event));
    scroller.addEventListener('pointermove', event => this.#handlePointerMove(event));
    scroller.addEventListener('pointerup', event => this.#handlePointerUp(event));
    scroller.addEventListener('pointercancel', event => this.#handlePointerUp(event));
    scroller.addEventListener('dblclick', event => this.#handleDoubleClick(event));
    scroller.addEventListener('copy', event => this.#handleCopy(event, false));
    scroller.addEventListener('cut', event => this.#handleCopy(event, true));
    scroller.addEventListener('paste', event => this.#handlePaste(event));

    editor.addEventListener('keydown', event => this.#handleEditorKeyDown(event));
    editor.addEventListener('blur', () => this.#editMode && this.commitEdit());

    new ResizeObserver(() => this.#scheduleRender()).observe(scroller);

    this.#updateColumnLayout();
  }

  /** @type {HTMLElement} */
  #content;

  /** @type {HTMLElement} */
  #body;

  /** @type {HTMLElement} */
  #rangeOverlay;

  /** @type {HTMLElement} */
  #activeOverlay;

  /** @type {HTMLInputElement} */
  #editor;

  /** @param {Event} event */
  #handleCellChange = event => {
    const { x, y: row } = /** @type {import('./workbook/private/CellChangeEvent.js').CellChangeEvent} */ (event);

    // Only re-render if the changed cell is currently rendered. Multiple changes are batched into a single frame.
    if ((row >= this.#firstRow && row < this.#firstRow + this.#pool.length) || row < this.#lockedRowCount) {
      this.#contentDirty = true;
      this.#scheduleRender();
    }

    if (row === this.#anchor.row && x === this.#anchor.column) {
      this.#dispatchSelectionChange();
    }
  };

  /** @param {Event} event */
  #handleColumnWidthChange = event => {
    this.#applyColumnWidth(/** @type {Event & { x: number }} */ (event).x);
    this.#updateColumnLayout();
    // Merged cells span multiple columns.
    this.#contentDirty = true;
    this.#scheduleRender();
  };

  #handleMergeChange = () => {
    this.#contentDirty = true;
    this.#scheduleRender();
    this.#anchor = this.#toMergeStart(this.#anchor);
    this.#updateSelection();
  };

  /**
   * Returns the first cell of the merged cell containing the cell, or the cell itself if not merged.
   *
   * @param {CellPosition} cell
   * @returns {CellPosition}
   */
  #toMergeStart(cell) {
    const merge = this.#worksheet?.getMerge(cell.column, cell.row);

    return merge ? { column: merge.x, row: cell.row } : cell;
  }

  /**
   * Returns the last column of the merged cell containing the cell, or the column of the cell if not merged.
   *
   * @param {CellPosition} cell
   */
  #mergeEndColumn({ column, row }) {
    const merge = this.#worksheet?.getMerge(column, row);

    return merge ? merge.x + merge.width - 1 : column;
  }

  #handleBorderChange = () => {
    this.#contentDirty = true;
    this.#scheduleRender();
  };

  /** @returns {Worksheet | undefined} */
  get worksheet() {
    return this.#worksheet;
  }

  /**
   * Shows a worksheet. Selection and scroll position are remembered per worksheet.
   *
   * @param {Worksheet | undefined} worksheet
   */
  setWorksheet(worksheet) {
    if (worksheet === this.#worksheet) {
      return;
    }

    this.#editMode && this.commitEdit();

    if (this.#worksheet) {
      this.#viewStates.set(this.#worksheet, {
        anchor: this.#anchor,
        focus: this.#focus,
        scrollLeft: this.element.scrollLeft,
        scrollTop: this.element.scrollTop
      });

      this.#worksheet.removeEventListener('cellchange', this.#handleCellChange);
      this.#worksheet.removeEventListener('borderchange', this.#handleBorderChange);
      this.#worksheet.removeEventListener('formattingchange', this.#handleBorderChange);
      this.#worksheet.removeEventListener('columnwidthchange', this.#handleColumnWidthChange);
      this.#worksheet.removeEventListener('mergechange', this.#handleMergeChange);
    }

    this.#worksheet = worksheet;

    const state = worksheet && this.#viewStates.get(worksheet);

    this.#anchor = state?.anchor || { column: 0, row: 0 };
    this.#focus = state?.focus || { column: 0, row: 0 };

    worksheet?.addEventListener('cellchange', this.#handleCellChange);
    worksheet?.addEventListener('borderchange', this.#handleBorderChange);
    worksheet?.addEventListener('formattingchange', this.#handleBorderChange);
    worksheet?.addEventListener('columnwidthchange', this.#handleColumnWidthChange);
    worksheet?.addEventListener('mergechange', this.#handleMergeChange);

    for (let column = 0; column < this.columnCount; column++) {
      this.#applyColumnWidth(column);
    }

    this.#updateColumnLayout();

    this.#clearScrollTarget();
    this.element.scrollLeft = state?.scrollLeft || 0;
    this.element.scrollTop = state?.scrollTop || 0;

    this.#contentDirty = true;
    this.#render();
    this.#updateSelection();
  }

  /** Whether lockable rows and columns are currently locked (read-only). */
  get locked() {
    return this.#locked;
  }

  set locked(locked) {
    locked = !!locked;

    if (locked === this.#locked) {
      return;
    }

    this.#locked = locked;
    this.element.classList.toggle('grid--locked', locked);
    locked && this.element.classList.remove('grid--resize-hover');

    if (locked && this.#editMode && !this.isEditable(this.#anchor.column, this.#anchor.row)) {
      this.cancelEdit();
    }
  }

  /**
   * Returns `true` if the cell can be edited by the user (i.e. not locked).
   *
   * @param {number} column
   * @param {number} row
   */
  isEditable(column, row) {
    return (
      !!this.#worksheet && !(this.#locked && (column < this.#lockedColumnCount || row < this.#lockedRowCount))
    );
  }

  /** @returns {CellPosition} */
  get activeCell() {
    return { ...this.#anchor };
  }

  /** @returns {CellRange} */
  get selection() {
    return {
      left: Math.min(this.#anchor.column, this.#focus.column),
      top: Math.min(this.#anchor.row, this.#focus.row),
      right: Math.max(this.#anchor.column, this.#focus.column),
      bottom: Math.max(this.#anchor.row, this.#focus.row)
    };
  }

  /**
   * Selects a cell or a range. The anchor is the active cell.
   *
   * @param {CellPosition} anchor
   * @param {CellPosition} [focus]
   */
  select(anchor, focus = anchor) {
    this.#editMode && this.commitEdit();

    this.#anchor = this.#toMergeStart(this.#clamp(anchor));
    this.#focus = this.#clamp(focus);

    this.#updateSelection();
    this.scrollIntoView(this.#focus);
  }

  focus() {
    this.element.focus({ preventScroll: true });
  }

  /**
   * Scrolls the grid (using the native scrollbar) so the cell is fully visible.
   *
   * Short distances (up to one viewport) are scrolled smoothly by the browser, unless `behavior` is `'instant'` or the
   * user prefers reduced motion. Long jumps (e.g. Ctrl+Down) are instant, same as Excel.
   *
   * @param {CellPosition} cell
   * @param {{ behavior?: 'auto' | 'instant' }} [options]
   */
  scrollIntoView({ column, row }, { behavior = 'auto' } = {}) {
    const { element, rowHeight, headerHeight, rowHeaderWidth } = this;
    const columnWidth = this.getColumnWidth(column);
    const { clientWidth, clientHeight, scrollLeft, scrollTop } = element;

    // While a smooth scroll is in progress, continue from where it will end, so repeated key presses add up.
    const from = this.#scrollTarget || { left: scrollLeft, top: scrollTop };
    const cellLeft = rowHeaderWidth + this.#columnOffset(column);
    const cellTop = headerHeight + row * rowHeight;
    const frozenWidth = this.#locked ? this.#columnOffset(this.#lockedColumnCount) : 0;
    const frozenHeight = this.#locked ? this.#lockedRowCount * rowHeight : 0;
    let { left, top } = from;

    // Frozen cells are always visible, so they never require scrolling in their frozen direction.
    if (!frozenWidth || column >= this.#lockedColumnCount) {
      if (cellLeft - rowHeaderWidth - frozenWidth < left) {
        left = cellLeft - rowHeaderWidth - frozenWidth;
      } else if (cellLeft + columnWidth > left + clientWidth) {
        left = cellLeft + columnWidth - clientWidth;
      }
    }

    if (!frozenHeight || row >= this.#lockedRowCount) {
      if (cellTop - headerHeight - frozenHeight < top) {
        top = cellTop - headerHeight - frozenHeight;
      } else if (cellTop + rowHeight > top + clientHeight) {
        top = cellTop + rowHeight - clientHeight;
      }
    }

    left = Math.max(0, Math.min(left, element.scrollWidth - clientWidth));
    top = Math.max(0, Math.min(top, element.scrollHeight - clientHeight));

    if (left === from.left && top === from.top) {
      return;
    }

    const smooth =
      behavior !== 'instant' &&
      !prefersReducedMotion.matches &&
      Math.abs(left - scrollLeft) <= clientWidth &&
      Math.abs(top - scrollTop) <= clientHeight;

    if (smooth) {
      this.#scrollTarget = { left, top };
      element.scrollTo({ left, top, behavior: 'smooth' });
      this.#armScrollEndFallback();
    } else {
      this.#clearScrollTarget();
      element.scrollTo({ left, top, behavior: 'instant' });
    }
  }

  #clearScrollTarget() {
    clearTimeout(this.#scrollEndTimeout);
    this.#scrollTarget = undefined;
  }

  /**
   * Fallback for browsers without the "scrollend" event: assume the smooth scroll ended (or was interrupted) when there
   * is no scroll event for a while.
   */
  #armScrollEndFallback() {
    if (!('onscrollend' in this.element)) {
      clearTimeout(this.#scrollEndTimeout);
      this.#scrollEndTimeout = setTimeout(() => (this.#scrollTarget = undefined), 150);
    }
  }

  #handleScroll() {
    this.#render();

    const target = this.#scrollTarget;

    if (target) {
      const { scrollLeft, scrollTop } = this.element;

      if (Math.abs(scrollLeft - target.left) < 1 && Math.abs(scrollTop - target.top) < 1) {
        this.#clearScrollTarget();
      } else {
        this.#armScrollEndFallback();
      }
    }
  }

  // #region Rendering

  #scheduleRender() {
    if (!this.#renderScheduled) {
      this.#renderScheduled = true;

      requestAnimationFrame(() => {
        this.#renderScheduled = false;
        this.#render();
      });
    }
  }

  #render() {
    const { element, rowHeight, headerHeight, rowCount } = this;
    const visibleRows = Math.ceil(Math.max(0, element.clientHeight - headerHeight) / rowHeight) + 1;
    const poolSize = Math.min(rowCount, visibleRows + Math.max(MIN_OVERSCAN, visibleRows) * 2);

    let poolGrew = false;

    while (this.#pool.length < poolSize) {
      const slot = this.#createRow();

      this.#body.insertBefore(slot.element, this.#rangeOverlay);
      this.#pool.push(slot);
      poolGrew = true;
    }

    // Pool size changed, row-to-slot mapping is no longer valid.
    poolGrew && this.#pool.forEach(slot => (slot.row = -1));

    const firstRow = Math.max(
      0,
      Math.min(
        rowCount - this.#pool.length,
        Math.floor(element.scrollTop / rowHeight) - Math.floor((this.#pool.length - visibleRows) / 2)
      )
    );
    const { length } = this.#pool;

    this.#firstRow = firstRow;

    const forceContent = this.#contentDirty;

    this.#contentDirty = false;

    for (let row = firstRow; row < firstRow + length; row++) {
      // Slots are recycled by row index, so only rows newly scrolled into view need to be updated.
      const slot = this.#pool[row % length];

      if (slot.row !== row) {
        slot.row = row;
        slot.element.style.top = `${row * rowHeight}px`;
        slot.element.setAttribute('aria-rowindex', String(row + 2));
        slot.header.textContent = String(row + 1);
        slot.element.classList.toggle('grid__row--lockable', row < this.#lockedRowCount);
        this.#renderRowSelection(slot);
        this.#renderRowContent(slot);
      } else if (forceContent) {
        this.#renderRowContent(slot);
      }
    }

    forceContent && this.#frozenRows.forEach(slot => this.#renderRowContent(slot));
  }

  /** @returns {{ element: HTMLElement; header: HTMLElement; cells: HTMLElement[]; row: number }} */
  #createRow() {
    const element = createElement('div', 'grid__row');
    const header = createElement('div', 'grid__row-header');
    const cells = [];

    element.setAttribute('role', 'row');
    header.setAttribute('role', 'rowheader');
    element.append(header);

    for (let column = 0; column < this.columnCount; column++) {
      const cell = createElement('div', 'grid__cell');

      cell.setAttribute('role', 'gridcell');
      cell.append(createElement('span', 'grid__cell-text'));
      cell.style.width = `${this.getColumnWidth(column)}px`;
      this.#initLockableColumn(cell, column);
      cells.push(cell);
      element.append(cell);
    }

    return { element, header, cells, row: -1 };
  }

  /**
   * Returns the left offset of a column, relative to the right of the row header.
   *
   * @param {number} column
   */
  #columnOffset(column) {
    let offset = 0;

    for (let index = 0; index < column; index++) {
      offset += this.getColumnWidth(index);
    }

    return offset;
  }

  /**
   * Returns the column at an offset relative to the right of the row header.
   *
   * @param {number} x
   */
  #columnAtOffset(x) {
    if (x < 0) {
      return -1;
    }

    for (let column = 0, offset = 0; column < this.columnCount; column++) {
      offset += this.getColumnWidth(column);

      if (x < offset) {
        return column;
      }
    }

    return this.columnCount;
  }

  /**
   * Returns the width of a column of the current worksheet.
   *
   * @param {number} column
   */
  getColumnWidth(column) {
    return this.#worksheet && column >= 0 && column < this.columnCount
      ? this.#worksheet.getColumnWidth(column)
      : this.columnWidth;
  }

  /**
   * Sets the width of a column of the current worksheet.
   *
   * @param {number} column
   * @param {number} width
   */
  setColumnWidth(column, width) {
    if (this.#worksheet && column >= 0 && column < this.columnCount) {
      this.#worksheet.setColumnWidth(column, Math.max(MIN_COLUMN_WIDTH, Math.round(width)));
    }
  }

  /**
   * Applies the width of a column to the column header and cells.
   *
   * @param {number} column
   */
  #applyColumnWidth(column) {
    const style = `${this.getColumnWidth(column)}px`;

    this.#columnHeaders[column].style.width = style;

    for (const { cells } of [...this.#pool, ...this.#frozenRows]) {
      cells[column].style.width = style;
    }
  }

  /**
   * Resizes columns to fit the content of all their cells and headers. All columns are set to the same width.
   *
   * @param {number[]} columns
   */
  autoFitColumns(columns) {
    const width = Math.max(...columns.map(column => this.#measureColumn(column)));

    columns.forEach(column => this.setColumnWidth(column, width));
  }

  /**
   * Resizes a column to fit the content of all its cells and its header.
   *
   * @param {number} column
   */
  autoFitColumn(column) {
    this.autoFitColumns([column]);
  }

  /**
   * Returns the width needed to fit the content of all cells in a column and its header.
   *
   * @param {number} column
   */
  #measureColumn(column) {
    const sample = this.#pool[0]?.cells[column] || this.#columnHeaders[column];
    const context = (this.#measureContext ||= document.createElement('canvas').getContext('2d'));
    const style = getComputedStyle(sample);
    // Padding plus right border.
    const extra = parseFloat(style.paddingLeft) + parseFloat(style.paddingRight) + 1;

    if (!context) {
      return this.getColumnWidth(column);
    }

    context.font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;

    let width = context.measureText(getColumnName(column)).width;

    if (this.#worksheet) {
      const formattings = this.#worksheet.getFormattings(column, 0, 1, this.rowCount);
      const values = this.#worksheet.getValues(column, 0, 1, this.rowCount);

      for (let row = 0; row < values.length; row++) {
        const merge = this.#worksheet.getMerge(column, row);
        // Content of a merged cell is in its first cell, and spans all merged columns.
        const value = merge ? this.#worksheet.getValue(merge.x, row) : values[row][0];
        const format = (merge ? this.#worksheet.getFormattings(merge.x, row, 1, 1)[0][0] : formattings[row][0]).format;
        let otherColumnsWidth = 0;

        if (merge) {
          for (let index = merge.x; index < merge.x + merge.width; index++) {
            index !== column && (otherColumnsWidth += this.getColumnWidth(index));
          }
        }

        const { text: formatted } = formatDisplayValue(value, format);
        // Includes decorations added by CSS: currency sign (with a gap) and parentheses, or percent sign.
        const text =
          typeof value === 'number' && format === 'currency' ? `$ (${formatted})` : typeof value === 'number' && format === 'percent' ? `${formatted}%` : formatted;

        if (text) {
          for (const line of text.split('\n')) {
            width = Math.max(width, context.measureText(line).width - otherColumnsWidth);
          }
        }
      }
    }

    return Math.ceil(width + extra + 2);
  }

  /** @type {CanvasRenderingContext2D | null | undefined} */
  #measureContext;

  /** Updates positions which depend on column widths. */
  #updateColumnLayout() {
    this.#content.style.width = `${this.rowHeaderWidth + this.#columnOffset(this.columnCount)}px`;

    for (let column = 0; column < this.#lockedColumnCount; column++) {
      const left = `${this.rowHeaderWidth + this.#columnOffset(column)}px`;

      this.#columnHeaders[column].style.left = left;

      for (const { cells } of [...this.#pool, ...this.#frozenRows]) {
        cells[column].style.left = left;
      }
    }

    this.#editMode && this.#positionEditor();
    this.#updateSelection();
  }

  /**
   * Marks a cell or column header in a locked column. While locked, it sticks to the left so the column is frozen.
   *
   * @param {HTMLElement} element
   * @param {number} column
   */
  #initLockableColumn(element, column) {
    if (column < this.#lockedColumnCount) {
      element.classList.add('grid__lockable');
      element.style.left = `${this.rowHeaderWidth + this.#columnOffset(column)}px`;
    }
  }

  /** @param {{ cells: HTMLElement[]; row: number }} slot */
  #renderRowContent({ cells, row }) {
    const worksheet = this.#worksheet;
    const merges = worksheet ? worksheet.getMerges(row) : [];
    let mergeIndex = 0;

    for (let column = 0; column < cells.length; column++) {
      const cell = cells[column];
      const merge = merges[mergeIndex]?.x <= column ? merges[mergeIndex] : undefined;
      const lastColumn = merge ? merge.x + merge.width - 1 : column;

      merge && column === lastColumn && mergeIndex++;

      // A merged cell is shown as its first cell spanning all merged columns, other cells are hidden.
      const hidden = !!merge && column !== merge.x;
      const width = `${merge ? this.#columnOffset(lastColumn + 1) - this.#columnOffset(column) : this.getColumnWidth(column)}px`;

      if (cell.style.display !== (hidden ? 'none' : '')) {
        cell.style.display = hidden ? 'none' : '';
      }

      if (cell.style.width !== width) {
        cell.style.width = width;
      }
      const value = worksheet ? worksheet.getValue(column, row) : '';
      const type = typeof value;
      const formatting = worksheet?.getFormattings(column, row, 1, 1)[0][0];
      // Number formats only apply to numbers, decorations (e.g. currency sign) are added by CSS.
      const format = (type === 'number' && formatting?.format) || '';
      const { negative, text } = formatDisplayValue(value, formatting?.format);

      // Text is in an inner element, so the highlight marker only covers the text, not the whole cell.
      const content = /** @type {HTMLElement} */ (cell.firstElementChild);

      if (content.textContent !== text) {
        content.textContent = text;
      }

      if (cell.dataset.type !== type) {
        cell.dataset.type = type;
      }

      if ((cell.dataset.format || '') !== format) {
        format ? (cell.dataset.format = format) : delete cell.dataset.format;
      }

      cell.classList.toggle('grid__cell--negative', negative);

      const decorations = [formatting?.underline && 'underline', formatting?.strikethrough && 'line-through']
        .filter(Boolean)
        .join(' ');
      const align = formatting?.align || '';

      cell.classList.toggle('grid__cell--bold', !!formatting?.bold);
      cell.classList.toggle('grid__cell--italic', !!formatting?.italic);

      if (cell.style.textDecorationLine !== decorations) {
        cell.style.textDecorationLine = decorations;
      }

      if ((cell.dataset.align || '') !== align) {
        align ? (cell.dataset.align = align) : delete cell.dataset.align;
      }

      const highlight = getHighlight(formatting?.highlight);
      const background = highlight?.background || '';
      const color = highlight?.color || '';

      // Compare with the last set value, as the browser normalizes colors.
      if ((cell.dataset.background || '') !== background) {
        background ? content.style.setProperty('--marker-color', background) : content.style.removeProperty('--marker-color');
        content.classList.toggle('grid__cell-text--marker', !!background);
        background ? (cell.dataset.background = background) : delete cell.dataset.background;
      }

      if ((cell.dataset.color || '') !== color) {
        cell.style.color = color;
        color ? (cell.dataset.color = color) : delete cell.dataset.color;
      }

      const shadow = worksheet ? this.#getBorderShadow(worksheet, column, row, lastColumn) : '';

      // Compare with the last set value, as the browser normalizes `style.boxShadow`.
      if ((cell.dataset.border || '') !== shadow) {
        cell.dataset.border = shadow;
        cell.style.boxShadow = shadow;
      }

      const corners = worksheet ? this.#getBorderCorners(worksheet, column, row, lastColumn) : '';

      if ((cell.dataset.borderCorners || '') !== corners) {
        cell.dataset.borderCorners = corners;
        corners ? cell.style.setProperty('--border-corners', corners) : cell.style.removeProperty('--border-corners');
        cell.toggleAttribute('data-border-corner', !!corners);
      }
    }
  }

  /**
   * Returns the inset box shadows drawing the borders of a cell. Each border is centered on the 1px grid line, which is
   * the last pixel of the cell on the left or above: that cell draws the grid line and 1px before it, the cell on the
   * right or below draws the remaining 1px.
   *
   * @param {Worksheet} worksheet
   * @param {number} column
   * @param {number} row
   * @param {number} [lastColumn] Last column of a merged cell.
   */
  #getBorderShadow(worksheet, column, row, lastColumn = column) {
    const inside = BORDER_WIDTH - BORDER_OUTSIDE_WIDTH;
    const shadows = [];

    hasVerticalEdge(worksheet, lastColumn, row) && shadows.push(`inset -${inside}px 0 ${BORDER_COLOR}`);
    hasHorizontalEdge(worksheet, column, row) && shadows.push(`inset 0 -${inside}px ${BORDER_COLOR}`);
    hasVerticalEdge(worksheet, column - 1, row) && shadows.push(`inset ${BORDER_OUTSIDE_WIDTH}px 0 ${BORDER_COLOR}`);
    hasHorizontalEdge(worksheet, column, row - 1) && shadows.push(`inset 0 ${BORDER_OUTSIDE_WIDTH}px ${BORDER_COLOR}`);

    return shadows.length ? [...shadows, 'var(--grid-line-shadow)'].join(', ') : '';
  }

  /**
   * Returns the background layers filling the parts of border junctions (where grid lines cross) inside a cell, so
   * borders meeting at a corner join without a notch. Empty if none.
   *
   * @param {Worksheet} worksheet
   * @param {number} column
   * @param {number} row
   * @param {number} [lastColumn] Last column of a merged cell.
   */
  #getBorderCorners(worksheet, column, row, lastColumn = column) {
    const inside = BORDER_WIDTH - BORDER_OUTSIDE_WIDTH;
    const outside = BORDER_OUTSIDE_WIDTH;
    const fill = `linear-gradient(${BORDER_COLOR}, ${BORDER_COLOR})`;
    const layers = [];

    hasJunction(worksheet, lastColumn, row) && layers.push(`${fill} right bottom / ${inside}px ${inside}px no-repeat`);
    hasJunction(worksheet, column - 1, row) && layers.push(`${fill} left bottom / ${outside}px ${inside}px no-repeat`);
    hasJunction(worksheet, lastColumn, row - 1) && layers.push(`${fill} right top / ${inside}px ${outside}px no-repeat`);
    hasJunction(worksheet, column - 1, row - 1) && layers.push(`${fill} left top / ${outside}px ${outside}px no-repeat`);

    return layers.join(', ');
  }

  /** @param {{ header: HTMLElement; row: number }} slot */
  #renderRowSelection({ header, row }) {
    const { top, bottom } = this.selection;

    header.classList.toggle('grid__row-header--selected', row >= top && row <= bottom);
  }

  #updateSelection() {
    const { rowHeight, rowHeaderWidth } = this;
    const { top, bottom } = this.selection;
    let { left, right } = this.selection;
    const rangeStyle = this.#rangeOverlay.style;
    const activeStyle = this.#activeOverlay.style;
    const anchorEnd = this.#mergeEndColumn(this.#anchor);

    // The highlighted range covers merged cells partially selected.
    if (this.#worksheet) {
      for (let row = top; row <= bottom; row++) {
        left = Math.min(left, this.#toMergeStart({ column: left, row }).column);
        right = Math.max(right, this.#mergeEndColumn({ column: right, row }));
      }
    }

    rangeStyle.left = `${rowHeaderWidth + this.#columnOffset(left)}px`;
    rangeStyle.top = `${top * rowHeight}px`;
    rangeStyle.width = `${this.#columnOffset(right + 1) - this.#columnOffset(left)}px`;
    rangeStyle.height = `${(bottom - top + 1) * rowHeight}px`;
    this.#rangeOverlay.hidden = left === this.#anchor.column && right === anchorEnd && top === bottom;

    activeStyle.left = `${rowHeaderWidth + this.#columnOffset(this.#anchor.column)}px`;
    activeStyle.width = `${this.#columnOffset(anchorEnd + 1) - this.#columnOffset(this.#anchor.column)}px`;
    activeStyle.top = `${this.#anchor.row * rowHeight}px`;

    this.#columnHeaders.forEach((header, column) =>
      header.classList.toggle('grid__column-header--selected', column >= left && column <= right)
    );

    this.#pool.forEach(slot => slot.row !== -1 && this.#renderRowSelection(slot));
    this.#frozenRows.forEach(slot => this.#renderRowSelection(slot));

    this.#dispatchSelectionChange();
  }

  #dispatchSelectionChange() {
    this.dispatchEvent(new Event('selectionchange'));
  }

  // #endregion

  // #region Keyboard

  /**
   * @param {CellPosition} position
   * @returns {CellPosition}
   */
  #clamp({ column, row }) {
    return {
      column: Math.max(0, Math.min(this.columnCount - 1, column)),
      row: Math.max(0, Math.min(this.rowCount - 1, row))
    };
  }

  /**
   * Moves the active cell, or extends the selection.
   *
   * @param {number} deltaColumn
   * @param {number} deltaRow
   * @param {boolean} extend
   */
  #move(deltaColumn, deltaRow, extend) {
    const from = extend ? this.#focus : this.#anchor;
    // Moving right skips the rest of a merged cell.
    const column = deltaColumn > 0 ? this.#mergeEndColumn(from) : from.column;
    const to = this.#clamp({ column: column + deltaColumn, row: from.row + deltaRow });

    extend ? this.select(this.#anchor, to) : this.select(to);
  }

  /**
   * Moves the active cell within the selection (Enter / Tab), wrapping around like Excel when a range is selected.
   *
   * @param {number} deltaColumn
   * @param {number} deltaRow
   */
  #moveWithinSelection(deltaColumn, deltaRow) {
    const { left, top, right, bottom } = this.selection;

    if (left === right && top === bottom) {
      return this.#move(deltaColumn, deltaRow, false);
    }

    let { column, row } = this.#anchor;

    // Moving right skips the rest of a merged cell.
    column = (deltaColumn > 0 ? this.#mergeEndColumn(this.#anchor) : column) + deltaColumn;
    row += deltaRow;

    if (column > right) {
      column = left;
      row = row + 1 > bottom ? top : row + 1;
    } else if (column < left) {
      column = right;
      row = row - 1 < top ? bottom : row - 1;
    } else if (row > bottom) {
      row = top;
      column = column + 1 > right ? left : column + 1;
    } else if (row < top) {
      row = bottom;
      column = column - 1 < left ? right : column - 1;
    }

    this.#selectKeepingScroll({ column, row }, this.#focus);
    this.scrollIntoView(this.#anchor);
  }

  /** @param {KeyboardEvent} event */
  #handleKeyDown(event) {
    if (event.target !== this.element || event.isComposing) {
      return;
    }

    const { key, shiftKey } = event;
    const ctrlKey = event.ctrlKey || event.metaKey;
    const frozenHeight = this.#locked ? this.#lockedRowCount * this.rowHeight : 0;
    const pageRows = Math.max(
      1,
      Math.floor((this.element.clientHeight - this.headerHeight - frozenHeight) / this.rowHeight) - 1
    );
    const { column, row } = shiftKey ? this.#focus : this.#anchor;

    /** @type {[number, number] | undefined} */
    let delta;

    switch (key) {
      case 'ArrowUp':
        delta = [0, ctrlKey ? -row : -1];
        break;

      case 'ArrowDown':
        delta = [0, ctrlKey ? this.rowCount - 1 - row : 1];
        break;

      case 'ArrowLeft':
        delta = [ctrlKey ? -column : -1, 0];
        break;

      case 'ArrowRight':
        delta = [ctrlKey ? this.columnCount - 1 - column : 1, 0];
        break;

      case 'PageUp':
        delta = [0, -pageRows];
        break;

      case 'PageDown':
        delta = [0, pageRows];
        break;

      case 'Home':
        delta = [-column, ctrlKey ? -row : 0];
        break;

      case 'End':
        delta = [this.columnCount - 1 - column, ctrlKey ? this.rowCount - 1 - row : 0];
        break;

      case 'Enter':
        event.preventDefault();
        this.#moveWithinSelection(0, shiftKey ? -1 : 1);

        return;

      case 'Tab':
        event.preventDefault();
        this.#moveWithinSelection(shiftKey ? -1 : 1, 0);

        return;

      case 'F2':
        event.preventDefault();
        this.startEdit();

        return;

      case 'Delete':
      case 'Backspace':
        event.preventDefault();
        this.clearSelection();

        return;

      default:
        if (ctrlKey && !event.altKey && key.toLowerCase() === 'a') {
          event.preventDefault();
          this.#selectKeepingScroll({ column: 0, row: 0 }, { column: this.columnCount - 1, row: this.rowCount - 1 });
        } else if (key.length === 1 && !ctrlKey && !event.altKey) {
          // Printable character: start editing, replacing the content (same as Excel).
          event.preventDefault();
          this.startEdit(key);
        }

        return;
    }

    event.preventDefault();
    this.#move(delta[0], delta[1], shiftKey);
  }

  // #endregion

  // #region Pointer

  /**
   * @param {PointerEvent | MouseEvent} event
   * @returns {CellPosition}
   */
  #cellFromPoint({ clientX, clientY }) {
    const { element, headerHeight, rowHeaderWidth, rowHeight } = this;
    const rect = this.#body.getBoundingClientRect();
    const scrollerRect = element.getBoundingClientRect();
    // Position relative to the top-left of the scrollable area (below the column header, right of the row header).
    const x = clientX - scrollerRect.left - element.clientLeft - rowHeaderWidth;
    const y = clientY - scrollerRect.top - element.clientTop - headerHeight;
    const frozen = this.#locked;

    return this.#clamp({
      column:
        frozen && x >= 0 && x < this.#columnOffset(this.#lockedColumnCount)
          ? this.#columnAtOffset(x)
          : this.#columnAtOffset(clientX - rect.left - rowHeaderWidth),
      row:
        frozen && y >= 0 && y < this.#lockedRowCount * rowHeight
          ? Math.floor(y / rowHeight)
          : Math.floor((clientY - rect.top) / rowHeight)
    });
  }

  /** @param {PointerEvent} event */
  #handlePointerDown(event) {
    const target = /** @type {HTMLElement} */ (event.target);

    if (event.button !== 0 || target === this.#editor) {
      return;
    }

    const scrollerRect = this.element.getBoundingClientRect();

    // Ignore clicks on the native scrollbars.
    if (
      event.clientX - scrollerRect.left >= this.element.clientLeft + this.element.clientWidth ||
      event.clientY - scrollerRect.top >= this.element.clientTop + this.element.clientHeight
    ) {
      return;
    }

    event.preventDefault();
    this.focus();

    const resizeColumn = this.#resizeColumnFromPoint(event);

    if (resizeColumn >= 0) {
      this.#dragging = 'resize';
      this.#resizing = {
        columns: this.#columnsToResize(resizeColumn),
        startX: event.clientX,
        startWidth: this.getColumnWidth(resizeColumn)
      };
      this.element.setPointerCapture(event.pointerId);

      return;
    }

    const cell = this.#cellFromPoint(event);
    const { shiftKey } = event;

    if (target.closest('.grid__corner')) {
      this.#selectKeepingScroll({ column: 0, row: 0 }, { column: this.columnCount - 1, row: this.rowCount - 1 });

      return;
    } else if (target.closest('.grid__column-header')) {
      const column = Number(/** @type {HTMLElement} */ (target.closest('.grid__column-header')).dataset.column);

      this.#dragging = 'column';
      shiftKey
        ? this.#selectKeepingScroll(this.#anchor, { column, row: this.rowCount - 1 })
        : this.#selectKeepingScroll({ column, row: 0 }, { column, row: this.rowCount - 1 });
    } else if (target.closest('.grid__row-header')) {
      this.#dragging = 'row';
      shiftKey
        ? this.#selectKeepingScroll(this.#anchor, { column: this.columnCount - 1, row: cell.row })
        : this.#selectKeepingScroll({ column: 0, row: cell.row }, { column: this.columnCount - 1, row: cell.row });
    } else {
      this.#dragging = 'cell';
      shiftKey ? this.select(this.#anchor, cell) : this.select(cell);
    }

    this.element.setPointerCapture(event.pointerId);
  }

  /**
   * Selects without scrolling to the focus cell, used for whole row/column selection.
   *
   * @param {CellPosition} anchor
   * @param {CellPosition} focus
   */
  #selectKeepingScroll(anchor, focus) {
    this.#editMode && this.commitEdit();
    this.#anchor = this.#toMergeStart(this.#clamp(anchor));
    this.#focus = this.#clamp(focus);
    this.#updateSelection();
  }

  /** @param {PointerEvent} event */
  #handlePointerMove(event) {
    if (!this.#dragging) {
      this.element.classList.toggle('grid--resize-hover', this.#resizeColumnFromPoint(event) >= 0);

      return;
    }

    if (this.#dragging === 'resize') {
      const { columns, startX, startWidth } = /** @type {NonNullable<typeof this.#resizing>} */ (this.#resizing);

      columns.forEach(column => this.setColumnWidth(column, startWidth + event.clientX - startX));

      return;
    }

    const cell = this.#cellFromPoint(event);

    if (this.#dragging === 'column') {
      this.#selectKeepingScroll(this.#anchor, { column: cell.column, row: this.rowCount - 1 });
    } else if (this.#dragging === 'row') {
      this.#selectKeepingScroll(this.#anchor, { column: this.columnCount - 1, row: cell.row });
    } else if (cell.column !== this.#focus.column || cell.row !== this.#focus.row) {
      this.select(this.#anchor, cell);
    }
  }

  /** @param {PointerEvent} event */
  #handlePointerUp(event) {
    if (this.#dragging) {
      this.#dragging = '';
      this.#resizing = undefined;
      this.element.hasPointerCapture(event.pointerId) && this.element.releasePointerCapture(event.pointerId);
    }
  }

  /**
   * Returns the column whose right border in the column header is under the pointer, or -1. Always -1 while locked.
   *
   * @param {PointerEvent | MouseEvent} event
   */
  #resizeColumnFromPoint({ clientX, clientY }) {
    if (this.#locked) {
      return -1;
    }

    const { element, rowHeaderWidth } = this;
    const scrollerRect = element.getBoundingClientRect();
    const y = clientY - scrollerRect.top - element.clientTop;

    if (y < 0 || y >= this.headerHeight) {
      return -1;
    }

    const viewportX = clientX - scrollerRect.left - element.clientLeft - rowHeaderWidth;

    if (viewportX < -RESIZE_HANDLE_SIZE || viewportX >= element.clientWidth - rowHeaderWidth) {
      return -1;
    }

    const lockedWidth = this.#locked ? this.#columnOffset(this.#lockedColumnCount) : 0;
    // Frozen columns do not scroll.
    const x = this.#locked && viewportX < lockedWidth + RESIZE_HANDLE_SIZE ? viewportX : viewportX + element.scrollLeft;
    let offset = 0;

    for (let column = 0; column < this.columnCount; column++) {
      offset += this.getColumnWidth(column);

      if (Math.abs(x - offset) <= RESIZE_HANDLE_SIZE) {
        return column;
      } else if (x < offset - RESIZE_HANDLE_SIZE) {
        break;
      }
    }

    return -1;
  }

  /**
   * Returns the columns to resize when resizing a column: all selected columns if the column is part of a whole-column selection, otherwise only the column.
   *
   * @param {number} column
   * @returns {number[]}
   */
  #columnsToResize(column) {
    const { left, top, right, bottom } = this.selection;

    if (top === 0 && bottom === this.rowCount - 1 && column >= left && column <= right) {
      return Array.from({ length: right - left + 1 }, (_, index) => left + index);
    }

    return [column];
  }

  /** @param {MouseEvent} event */
  #handleDoubleClick(event) {
    const target = /** @type {HTMLElement} */ (event.target);
    const resizeColumn = this.#resizeColumnFromPoint(event);

    if (resizeColumn >= 0) {
      this.autoFitColumns(this.#columnsToResize(resizeColumn));

      return;
    }

    // With pointer capture, the target may be the grid itself, so hit-test using coordinates.
    if (
      target !== this.#editor &&
      !target.closest('.grid__header, .grid__row-header') &&
      event.clientY >= this.#body.getBoundingClientRect().top &&
      event.clientX >= this.element.getBoundingClientRect().left + this.rowHeaderWidth
    ) {
      this.startEdit();
    }
  }

  // #endregion

  // #region Editing

  /**
   * Opens the in-cell editor on the active cell.
   *
   * @param {string} [initialText] When specified, replaces the cell content ("enter" mode). Otherwise, edits the current content ("edit" mode).
   * @returns {boolean} `false` if the cell is locked.
   */
  startEdit(initialText) {
    const { column, row } = this.#anchor;

    if (!this.#worksheet || !this.isEditable(column, row)) {
      this.dispatchEvent(new CustomEvent('lockedit', { detail: { column, row } }));

      return false;
    }

    this.scrollIntoView(this.#anchor);

    const editor = this.#editor;
    const value = this.#worksheet.getValue(column, row);

    this.#positionEditor();
    editor.value = initialText ?? formatValue(value);
    editor.hidden = false;
    editor.focus({ preventScroll: true });
    editor.setSelectionRange(editor.value.length, editor.value.length);

    this.#editMode = initialText === undefined ? 'edit' : 'enter';

    return true;
  }

  /** Commits the editor value to the active cell. */
  #positionEditor() {
    const { column, row } = this.#anchor;
    const style = this.#editor.style;

    style.left = `${this.rowHeaderWidth + this.#columnOffset(column)}px`;
    style.top = `${row * this.rowHeight}px`;
    style.width = `${this.#columnOffset(this.#mergeEndColumn(this.#anchor) + 1) - this.#columnOffset(column)}px`;
    style.height = `${this.rowHeight}px`;
  }

  commitEdit() {
    if (!this.#editMode) {
      return;
    }

    const { column, row } = this.#anchor;
    const text = this.#editor.value;

    this.#closeEditor();

    if (this.#worksheet && this.isEditable(column, row)) {
      // Leading apostrophe forces the value to be text, same as Excel.
      // Leading "=" evaluates a simple arithmetic formula and stores only the result; invalid formulas are kept as text.
      const result = text.startsWith('=') ? calculate(text.slice(1)) : undefined;

      this.#worksheet.setValue(
        column,
        row,
        text.startsWith("'") ? text.slice(1) : result ?? inferValue(text)
      );
    }
  }

  /** Closes the editor without saving. */
  cancelEdit() {
    this.#editMode && this.#closeEditor();
  }

  #closeEditor() {
    const hadFocus = document.activeElement === this.#editor;

    this.#editMode = '';
    this.#editor.hidden = true;

    hadFocus && this.focus();
  }

  /** @param {KeyboardEvent} event */
  #handleEditorKeyDown(event) {
    if (event.isComposing) {
      return;
    }

    const { key, shiftKey } = event;

    if (key === 'Escape') {
      event.preventDefault();
      this.cancelEdit();
    } else if (key === 'Enter' && !event.altKey) {
      event.preventDefault();
      this.commitEdit();
      this.#moveWithinSelection(0, shiftKey ? -1 : 1);
    } else if (key === 'Tab') {
      event.preventDefault();
      this.commitEdit();
      this.#moveWithinSelection(shiftKey ? -1 : 1, 0);
    } else if (key === 'F2') {
      event.preventDefault();
      this.#editMode = this.#editMode === 'edit' ? 'enter' : 'edit';
    } else if (this.#editMode === 'enter' && key.startsWith('Arrow')) {
      event.preventDefault();
      this.commitEdit();
      this.#move(
        key === 'ArrowLeft' ? -1 : key === 'ArrowRight' ? 1 : 0,
        key === 'ArrowUp' ? -1 : key === 'ArrowDown' ? 1 : 0,
        false
      );
    }

    // Keep keystrokes inside the editor from reaching the grid.
    event.stopPropagation();
  }

  /** Clears all editable cells in the selection. */
  clearSelection() {
    const worksheet = this.#worksheet;

    if (!worksheet) {
      return;
    }

    const { left, top, right, bottom } = this.selection;
    let skipped = false;

    for (let row = top; row <= bottom; row++) {
      for (let column = left; column <= right; column++) {
        if (this.isEditable(column, row)) {
          worksheet.setValue(column, row, '');
        } else {
          skipped = true;
        }
      }
    }

    skipped && this.dispatchEvent(new CustomEvent('lockedit', { detail: this.#anchor }));
  }

  // #endregion

  // #region Clipboard

  /** @returns {CellValue[][]} */
  #getSelectedValues() {
    const worksheet = this.#worksheet;
    const { left, top, right, bottom } = this.selection;

    /** @type {CellValue[][]} */
    const values = [];

    if (worksheet) {
      for (let row = top; row <= bottom; row++) {
        const rowValues = [];

        for (let column = left; column <= right; column++) {
          rowValues.push(worksheet.getValue(column, row) ?? '');
        }

        values.push(rowValues);
      }
    }

    return values;
  }

  /**
   * @param {ClipboardEvent} event
   * @param {boolean} cut
   */
  #handleCopy(event, cut) {
    if (event.target !== this.element || !this.#worksheet || !event.clipboardData) {
      return;
    }

    event.preventDefault();

    const values = this.#getSelectedValues();
    const { left, top, right, bottom } = this.selection;
    // Formatting (excluding borders) is embedded in the HTML flavor only, so plain text editors get plain values.
    const formattings = this.#worksheet.getFormattings(left, top, right - left + 1, bottom - top + 1);

    // Excel and Excel Web App understand both plain text (tab-separated) and HTML tables.
    event.clipboardData.setData('text/plain', `${stringifyTSV(values)}\r\n`);
    event.clipboardData.setData(
      'text/html',
      `<table data-formattings="${escapeHTML(JSON.stringify(formattings))}">${values
        .map(
          row =>
            `<tr>${row
              .map(value =>
                typeof value === 'string'
                  ? `<td style="mso-number-format:'\\@'">${escapeHTML(value).replace(/\r?\n/g, '<br style="mso-data-placement:same-cell" />')}</td>`
                  : `<td>${escapeHTML(formatValue(value))}</td>`
              )
              .join('')}</tr>`
        )
        .join('')}</table>`
    );

    cut && this.clearSelection();
  }

  /** @param {ClipboardEvent} event */
  #handlePaste(event) {
    const worksheet = this.#worksheet;

    if (event.target !== this.element || !worksheet || !event.clipboardData) {
      return;
    }

    event.preventDefault();

    const text = event.clipboardData.getData('text/plain');

    if (!text) {
      return;
    }

    let values = parseTSV(text);
    const formattings = parseClipboardFormattings(event.clipboardData.getData('text/html'));

    if (!values.length) {
      values = [['']];
    }

    const dataRows = values.length;
    const dataColumns = Math.max(...values.map(row => row.length));
    const { left, top, right, bottom } = this.selection;
    const selectedRows = bottom - top + 1;
    const selectedColumns = right - left + 1;

    // Same as Excel: if the selection is a multiple of the copied data, the data is repeated to fill the selection.
    const tile = selectedRows % dataRows === 0 && selectedColumns % dataColumns === 0;
    const pasteRows = Math.min(tile ? selectedRows : dataRows, this.rowCount - top);
    const pasteColumns = Math.min(tile ? selectedColumns : dataColumns, this.columnCount - left);
    let skipped = false;

    for (let rowOffset = 0; rowOffset < pasteRows; rowOffset++) {
      const rowValues = values[rowOffset % dataRows];

      for (let columnOffset = 0; columnOffset < pasteColumns; columnOffset++) {
        const column = left + columnOffset;
        const row = top + rowOffset;

        if (this.isEditable(column, row)) {
          worksheet.setValue(column, row, rowValues[columnOffset % dataColumns] ?? '');

          if (formattings) {
            const formattingRow = formattings[rowOffset % dataRows];

            worksheet.setFormattings([[formattingRow?.[columnOffset % (formattingRow?.length || 1)]]], column, row);
          }
        } else {
          skipped = true;
        }
      }
    }

    const anchor = { column: left, row: top };

    this.#selectKeepingScroll(anchor, { column: left + pasteColumns - 1, row: top + pasteRows - 1 });

    skipped && this.dispatchEvent(new CustomEvent('lockedit', { detail: anchor }));
  }

  // #endregion
}
