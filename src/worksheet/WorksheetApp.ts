import { calculate } from '../calculator.ts';
import { formatValue, inferValue, parseTSV, stringifyTSV } from '../csv.ts';
import { parseHighlight } from './highlight.ts';
import { formatDisplayValue } from './numberFormats.ts';
import { getColumnName } from './getColumnName.ts';
import type { CellFormatting, CellValue } from './Worksheet.ts';
import { Worksheet } from './Worksheet.ts';
import type { WorksheetChangeEvent } from './WorksheetChangeEvent.ts';
import { type CellPosition, type CellRange, WorksheetSelection } from './WorksheetSelection.ts';

type RowSlot = { element: HTMLElement; header: HTMLElement; cells: HTMLElement[]; row: number };

/**
 * Minimum number of extra rows rendered above and below the visible area. Otherwise, about one viewport of extra rows is
 * rendered on each side, so fast or momentum scrolling (which may be composited ahead of the main thread) does not show
 * blank rows.
 */
const MIN_OVERSCAN = 4;

const prefersReducedMotion =
  typeof matchMedia === 'function' ? matchMedia('(prefers-reduced-motion: reduce)') : { matches: false };

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
 */
function getBorderOrNone(worksheet: Worksheet, column: number, row: number) {
  return column >= 0 && row >= 0 && column < worksheet.columnCount && row < worksheet.rowCount
    ? worksheet.getBorder(column, row)
    : undefined;
}

/**
 * Returns `true` if there is a border on the grid line on the right of the cell (column may be -1).
 */
function hasVerticalEdge(worksheet: Worksheet, column: number, row: number) {
  return !!(getBorderOrNone(worksheet, column, row)?.right || getBorderOrNone(worksheet, column + 1, row)?.left);
}

/**
 * Returns `true` if there is a border on the grid line below the cell (row may be -1).
 */
function hasHorizontalEdge(worksheet: Worksheet, column: number, row: number) {
  return !!(getBorderOrNone(worksheet, column, row)?.bottom || getBorderOrNone(worksheet, column, row + 1)?.top);
}

/**
 * Returns `true` if any border ends at or passes through the bottom-right corner of the cell.
 */
function hasJunction(worksheet: Worksheet, column: number, row: number) {
  return (
    hasVerticalEdge(worksheet, column, row) ||
    hasVerticalEdge(worksheet, column, row + 1) ||
    hasHorizontalEdge(worksheet, column, row) ||
    hasHorizontalEdge(worksheet, column + 1, row)
  );
}

function createElement(tagName: string, className: string): HTMLElement {
  const element = document.createElement(tagName);

  element.className = className;

  return element;
}

/**
 * Returns formattings embedded by this app in clipboard HTML, or `undefined` if the HTML is not from this app.
 */
function parseClipboardFormattings(html: string): Partial<CellFormatting>[][] | undefined {
  if (!html) {
    return;
  }

  try {
    const json = new DOMParser()
      .parseFromString(html, 'text/html')
      .querySelector('table[data-formattings]')
      ?.getAttribute('data-formattings');
    const formattings = json && JSON.parse(json);

    return Array.isArray(formattings) && formattings.every(Array.isArray) ? formattings : undefined;
  } catch {
    return;
  }
}

function escapeHTML(text: string): string {
  return text.replace(
    /[&<>"']/g,
    char =>
      (({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }) as Record<string, string>)[char] ||
      char
  );
}

/**
 * `<worksheet-app>`, a virtualized spreadsheet grid showing the `Worksheet` it owns. Only the visible rows (plus a
 * few extra) are in the DOM, while the scrolling itself is handled by the browser using a native scrollbar on the element
 * sized to the full worksheet.
 *
 * Public members:
 * - `worksheet`: a newly constructed `Worksheet` owned by the element. The selection (`worksheet.selection`) and the
 *   undo history (`worksheet.undo()`, `worksheet.redo()`) are on the worksheet.
 *
 * Attributes:
 * - `locked`: lockable rows and columns are read-only and frozen.
 * - `locked-column-count` and `locked-row-count`: number of lockable leading columns and rows, defaults to 0.
 *
 * Events: `lockedit` (`CustomEvent` with the `{ column, row }` of the locked cell).
 */
class WorksheetApp extends HTMLElement {
  static observedAttributes = ['locked', 'locked-column-count', 'locked-row-count'];

  readonly #worksheet = new Worksheet();

  #initialized = false;

  #columnWidth = 0;
  #rowHeight = 0;
  #headerHeight = 0;
  #rowHeaderWidth = 0;
  #columnCount = 0;
  #rowCount = 0;

  /** Cell the active cell was entered from, used when moving out of a merged cell. */
  #preferred: CellPosition = { column: 0, row: 0 };

  /** Cell the selection was extended to (before expanding to merged cells). */
  #focus: CellPosition = { column: 0, row: 0 };

  /** Whether the element itself is setting the selection of the worksheet. */
  #settingSelection = false;

  /** Cell being edited by the editor. */
  #editCell: CellPosition = { column: 0, row: 0 };

  #locked = false;

  /** Number of leading columns that are read-only and frozen while locked. */
  #lockedColumnCount = 0;

  /** Number of leading rows that are read-only and frozen while locked. */
  #lockedRowCount = 0;

  /** Copies of the locked rows, shown in a sticky pane while locked. */
  #frozenRows: RowSlot[] = [];

  #pool: RowSlot[] = [];

  #columnHeaders: HTMLElement[] = [];

  /** Whether the editor is open, and whether arrow keys should commit ("enter" mode) or move the caret ("edit" mode). */
  #editMode = '' as '' | 'enter' | 'edit';

  #renderScheduled = false;

  /** Where an in-progress smooth scroll will end. */
  #scrollTarget: { left: number; top: number } | undefined;

  #scrollEndTimeout: ReturnType<typeof setTimeout> | undefined;
  #firstRow = 0;
  #contentDirty = false;

  #dragging: 'cell' | 'row' | 'column' | 'resize' | '' = '';

  /** Columns being resized and where the drag started. */
  #resizing: { columns: number[]; startX: number; startWidth: number } | undefined;

  connectedCallback() {
    if (!this.#initialized) {
      this.#initialized = true;
      this.#init();
    }
  }

  attributeChangedCallback(name: string, _oldValue: string | null, value: string | null) {
    if (!this.#initialized) {
      return;
    }

    if (name === 'locked') {
      this.#setLocked(value !== null);
    } else {
      this.#updateLockedCounts();
    }
  }

  /** Reads `locked-column-count` and `locked-row-count`, and updates the lockable columns and rows. */
  #updateLockedCounts() {
    const lockedColumnCount = Math.max(
      0,
      Math.min(this.#columnCount, Math.floor(Number(this.getAttribute('locked-column-count'))) || 0)
    );
    const lockedRowCount = Math.max(
      0,
      Math.min(this.#rowCount, Math.floor(Number(this.getAttribute('locked-row-count'))) || 0)
    );

    if (lockedColumnCount === this.#lockedColumnCount && lockedRowCount === this.#lockedRowCount) {
      return;
    }

    this.#lockedColumnCount = lockedColumnCount;
    this.#lockedRowCount = lockedRowCount;
    this.style.setProperty('--locked-row-count', String(lockedRowCount));

    for (let column = 0; column < this.#columnCount; column++) {
      for (const element of [this.#columnHeaders[column]!, ...this.#pool.map(({ cells }) => cells[column]!)]) {
        this.#initLockableColumn(element, column);
      }
    }

    for (const slot of this.#pool) {
      slot.element.classList.toggle('grid__row--lockable', slot.row >= 0 && slot.row < lockedRowCount);
    }

    // Recreate the copies of the locked rows shown in the sticky pane.
    this.#frozenRows = [];

    for (let row = 0; row < lockedRowCount; row++) {
      const slot = this.#createRow();

      slot.element.classList.add('grid__row--lockable');
      slot.element.setAttribute('aria-hidden', 'true');
      slot.header.textContent = String(row + 1);
      slot.row = row;
      this.#frozenRows.push(slot);
      this.#renderRowSelection(slot);
    }

    this.#frozen.replaceChildren(...this.#frozenRows.map(({ element }) => element));

    if (
      this.#locked &&
      this.#editMode &&
      !this.#isEditable(this.#worksheet.selection.activeCell.column, this.#worksheet.selection.activeCell.row)
    ) {
      this.#cancelEdit();
    }

    this.#updateColumnLayout();
    this.#contentDirty = true;
    this.#scheduleRender();
  }

  #init() {
    const columnWidth = 100;
    const rowHeight = 24;
    const rowHeaderWidth = 56;
    const { columnCount, rowCount } = this.#worksheet;

    this.#columnWidth = columnWidth;
    this.#rowHeight = rowHeight;
    this.#headerHeight = rowHeight;
    this.#rowHeaderWidth = rowHeaderWidth;
    this.#columnCount = columnCount;
    this.#rowCount = rowCount;

    const scroller = this;

    scroller.classList.add('grid');

    scroller.tabIndex = 0;
    scroller.setAttribute('role', 'grid');
    scroller.setAttribute('aria-rowcount', String(rowCount + 1));
    scroller.setAttribute('aria-colcount', String(columnCount + 1));
    scroller.style.setProperty('--column-width', `${columnWidth}px`);
    scroller.style.setProperty('--row-height', `${rowHeight}px`);
    scroller.style.setProperty('--row-header-width', `${rowHeaderWidth}px`);
    scroller.style.setProperty('--locked-row-count', '0');

    const content = createElement('div', 'grid__content');

    this.#content = content;

    const header = createElement('div', 'grid__header');

    header.setAttribute('role', 'row');
    header.append(createElement('div', 'grid__corner'));

    for (let column = 0; column < columnCount; column++) {
      const columnHeader = createElement('div', 'grid__column-header');

      columnHeader.textContent = getColumnName(column);
      columnHeader.dataset['column'] = String(column);
      columnHeader.style.width = `${columnWidth}px`;
      columnHeader.setAttribute('role', 'columnheader');
      this.#initLockableColumn(columnHeader, column);
      this.#columnHeaders.push(columnHeader);
      header.append(columnHeader);
    }

    // Locked rows are frozen by showing copies of them in a sticky pane below the header, while the body is shifted up
    // underneath it so all other coordinates stay the same.
    const frozen = (this.#frozen = createElement('div', 'grid__frozen'));

    const body = (this.#body = createElement('div', 'grid__body'));

    body.style.height = `${rowCount * rowHeight}px`;

    this.#rangeOverlay = createElement('div', 'grid__range');
    this.#activeOverlay = createElement('div', 'grid__active');

    const editor = (this.#editor = createElement('input', 'grid__editor') as HTMLInputElement);

    editor.type = 'text';
    editor.spellcheck = false;
    editor.autocomplete = 'off';
    editor.hidden = true;

    body.append(this.#rangeOverlay, this.#activeOverlay, editor);
    content.append(header, frozen, body);
    scroller.append(content);

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
    editor.addEventListener('blur', () => this.#editMode && this.#commitEdit());

    new ResizeObserver(() => this.#scheduleRender()).observe(scroller);

    this.#attachWorksheet();
    this.#updateLockedCounts();
    this.#setLocked(this.hasAttribute('locked'));
  }

  #content!: HTMLElement;

  /** Sticky pane showing copies of the locked rows. */
  #frozen!: HTMLElement;

  #body!: HTMLElement;

  #rangeOverlay!: HTMLElement;

  #activeOverlay!: HTMLElement;

  #editor!: HTMLInputElement;

  /**
   * Re-renders what a transaction changed.
   */
  #handleChange = (event: Event) => {
    const { changes, load, undo } = event as WorksheetChangeEvent;

    // Undo or redo discards the pending edit.
    !load && !undo && this.#cancelEdit();

    for (const change of changes) {
      if (change.type === 'value') {
        const { x, y: row } = change;
        const merge = row < this.#firstRow ? this.#worksheet.getMerges(x, row, 1, 1)[0] : undefined;

        // Only re-render if the changed cell is currently rendered (or a merged cell extending into the rendered rows).
        if (
          (row >= this.#firstRow && row < this.#firstRow + this.#pool.length) ||
          row < this.#lockedRowCount ||
          (merge && merge.y + merge.height > this.#firstRow)
        ) {
          this.#contentDirty = true;
        }
      } else {
        this.#contentDirty = true;
      }
    }

    this.#scheduleRender();
  };

  /**
   * Shows the selection. If it was set from outside (e.g. by undo), also commits the pending edit and scrolls to it.
   */
  #handleSelectionChange = () => {
    this.#updateSelection();

    if (!this.#settingSelection) {
      const { activeCell, focus } = this.#worksheet.selection;

      this.#editMode && this.#commitEdit();
      this.#focus = focus;
      this.#preferred = activeCell;
      this.#scrollIntoView(focus);
    }
  };

  #handleColumnWidthChange = (event: Event) => {
    this.#applyColumnWidth((event as CustomEvent<{ x: number }>).detail.x);
    this.#updateColumnLayout();
    // Merged cells span multiple columns.
    this.#contentDirty = true;
    this.#scheduleRender();
  };

  /**
   * Returns the top-left cell of the merged cell containing the cell, or the cell itself if not merged.
   */
  #toMergeStart(cell: CellPosition): CellPosition {
    const merge = this.#worksheet.getMerges(cell.column, cell.row, 1, 1)[0];

    return merge ? { column: merge.x, row: merge.y } : cell;
  }

  /**
   * Returns the bottom-right cell of the merged cell containing the cell, or the cell itself if not merged.
   */
  #toMergeEnd(cell: CellPosition): CellPosition {
    const merge = this.#worksheet.getMerges(cell.column, cell.row, 1, 1)[0];

    return merge ? { column: merge.x + merge.width - 1, row: merge.y + merge.height - 1 } : cell;
  }

  get worksheet(): Worksheet {
    return this.#worksheet;
  }

  #attachWorksheet() {
    const worksheet = this.#worksheet;

    worksheet.addEventListener('change', this.#handleChange);
    worksheet.addEventListener('selectionchange', this.#handleSelectionChange);
    worksheet.addEventListener('columnwidthchange', this.#handleColumnWidthChange);

    for (let column = 0; column < this.#columnCount; column++) {
      this.#applyColumnWidth(column);
    }

    this.#updateColumnLayout();

    this.#contentDirty = true;
    this.#render();
    this.#updateSelection();
  }

  /**
   * Sets whether lockable rows and columns are locked (read-only).
   */
  #setLocked(locked: boolean) {
    if (locked === this.#locked) {
      return;
    }

    this.#locked = locked;
    this.classList.toggle('grid--locked', locked);
    locked && this.classList.remove('grid--resize-hover');

    if (
      locked &&
      this.#editMode &&
      !this.#isEditable(this.#worksheet.selection.activeCell.column, this.#worksheet.selection.activeCell.row)
    ) {
      this.#cancelEdit();
    }
  }

  /**
   * Returns `true` if the cell can be edited by the user (i.e. not locked).
   */
  #isEditable(column: number, row: number) {
    return !(this.#locked && (column < this.#lockedColumnCount || row < this.#lockedRowCount));
  }

  /**
   * Selects from an anchor to a focus. The worksheet snaps the anchor to the start of its merged cell and expands the
   * range to cover merged cells.
   */
  #setSelection(anchor: CellPosition, focus: CellPosition = anchor) {
    const activeCell = this.#clamp(anchor);

    this.#focus = this.#clamp(focus);

    this.#setSelectionObject(
      new WorksheetSelection(activeCell, {
        left: Math.min(activeCell.column, this.#focus.column),
        top: Math.min(activeCell.row, this.#focus.row),
        right: Math.max(activeCell.column, this.#focus.column),
        bottom: Math.max(activeCell.row, this.#focus.row)
      })
    );
  }

  #setSelectionObject(selection: WorksheetSelection) {
    this.#settingSelection = true;

    try {
      this.#worksheet.selection = selection;
    } finally {
      this.#settingSelection = false;
    }
  }

  /**
   * Returns the range between two cells, expanded to cover all merged cells partially in it.
   */
  #rangeOf(anchor: CellPosition, focus: CellPosition): CellRange {
    let top = Math.min(anchor.row, focus.row);
    let bottom = Math.max(anchor.row, focus.row);
    let left = Math.min(anchor.column, focus.column);
    let right = Math.max(anchor.column, focus.column);

    // The selection covers all merged cells partially selected. Expanding may cover more merged cells, so it repeats
    // until nothing changes.
    for (let changed = true; changed;) {
      changed = false;

      for (const { x, y, width, height } of this.#worksheet.getMerges(left, top, right - left + 1, bottom - top + 1)) {
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

  /**
   * Selects a cell or a range. The anchor is the active cell.
   */
  #select(anchor: CellPosition, focus: CellPosition = anchor) {
    this.#editMode && this.#commitEdit();

    this.#setSelection(anchor, focus);
    this.#preferred = this.#clamp(anchor);

    this.#scrollIntoView(this.#focus);
  }

  /**
   * Scrolls the grid (using the native scrollbar) so the cell is fully visible.
   *
   * Short distances (up to one viewport) are scrolled smoothly by the browser, unless `behavior` is `'instant'` or the
   * user prefers reduced motion. Long jumps (e.g. Ctrl+Down) are instant, same as Excel.
   */
  #scrollIntoView({ column, row }: CellPosition, { behavior = 'auto' }: { behavior?: 'auto' | 'instant' } = {}) {
    const element = this;
    const rowHeight = this.#rowHeight;
    const headerHeight = this.#headerHeight;
    const rowHeaderWidth = this.#rowHeaderWidth;
    const columnWidth = this.#getColumnWidth(column);
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
    if (!('onscrollend' in (this as object))) {
      clearTimeout(this.#scrollEndTimeout);
      this.#scrollEndTimeout = setTimeout(() => (this.#scrollTarget = undefined), 150);
    }
  }

  #handleScroll() {
    this.#render();

    const target = this.#scrollTarget;

    if (target) {
      const { scrollLeft, scrollTop } = this;

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
    const element = this;
    const rowHeight = this.#rowHeight;
    const headerHeight = this.#headerHeight;
    const rowCount = this.#rowCount;
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

    // Tall merged cells are drawn by their first rendered row, which changes with the first row.
    const forceContent =
      this.#contentDirty ||
      (firstRow !== this.#firstRow && this.#worksheet.getMerges().some(({ height }) => height > 1));

    this.#firstRow = firstRow;

    this.#contentDirty = false;

    for (let row = firstRow; row < firstRow + length; row++) {
      // Slots are recycled by row index, so only rows newly scrolled into view need to be updated.
      const slot = this.#pool[row % length]!;

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

  #createRow(): RowSlot {
    const element = createElement('div', 'grid__row');
    const header = createElement('div', 'grid__row-header');
    const cells: HTMLElement[] = [];

    element.setAttribute('role', 'row');
    header.setAttribute('role', 'rowheader');
    element.append(header);

    for (let column = 0; column < this.#columnCount; column++) {
      const cell = createElement('div', 'grid__cell');

      cell.setAttribute('role', 'gridcell');
      cell.append(createElement('span', 'grid__cell-text'));
      cell.style.width = `${this.#getColumnWidth(column)}px`;
      this.#initLockableColumn(cell, column);
      cells.push(cell);
      element.append(cell);
    }

    return { element, header, cells, row: -1 };
  }

  /**
   * Returns the left offset of a column, relative to the right of the row header.
   */
  #columnOffset(column: number) {
    let offset = 0;

    for (let index = 0; index < column; index++) {
      offset += this.#getColumnWidth(index);
    }

    return offset;
  }

  /**
   * Returns the column at an offset relative to the right of the row header.
   */
  #columnAtOffset(x: number) {
    if (x < 0) {
      return -1;
    }

    for (let column = 0, offset = 0; column < this.#columnCount; column++) {
      offset += this.#getColumnWidth(column);

      if (x < offset) {
        return column;
      }
    }

    return this.#columnCount;
  }

  /**
   * Returns the width of a column of the current worksheet.
   */
  #getColumnWidth(column: number) {
    return column >= 0 && column < this.#columnCount ? this.#worksheet.getColumnWidth(column) : this.#columnWidth;
  }

  /**
   * Sets the width of a column of the current worksheet.
   */
  #setColumnWidth(column: number, width: number) {
    if (column >= 0 && column < this.#columnCount) {
      this.#worksheet.setColumnWidth(column, Math.max(MIN_COLUMN_WIDTH, Math.round(width)));
    }
  }

  /**
   * Applies the width of a column to the column header and cells.
   */
  #applyColumnWidth(column: number) {
    const style = `${this.#getColumnWidth(column)}px`;

    this.#columnHeaders[column]!.style.width = style;

    for (const { cells } of [...this.#pool, ...this.#frozenRows]) {
      cells[column]!.style.width = style;
    }
  }

  /**
   * Resizes columns to fit the content of all their cells and headers. All columns are set to the same width.
   */
  #autoFitColumns(columns: number[]) {
    const width = Math.max(...columns.map(column => this.#measureColumn(column)));

    columns.forEach(column => this.#setColumnWidth(column, width));
  }

  /**
   * Returns the width needed to fit the content of all cells in a column and its header.
   */
  #measureColumn(column: number) {
    const sample = this.#pool[0]?.cells[column] || this.#columnHeaders[column]!;
    const context = (this.#measureContext ||= document.createElement('canvas').getContext('2d'));
    const style = getComputedStyle(sample);
    // Padding plus right border.
    const extra = parseFloat(style.paddingLeft) + parseFloat(style.paddingRight) + 1;

    if (!context) {
      return this.#getColumnWidth(column);
    }

    context.font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;

    let width = context.measureText(getColumnName(column)).width;

    const formattings = this.#worksheet.getFormattings(column, 0, 1, this.#rowCount);
    const values = this.#worksheet.getValues(column, 0, 1, this.#rowCount);

    for (let row = 0; row < values.length; row++) {
      const merge = this.#worksheet.getMerges(column, row, 1, 1)[0];

      // Content of a merged cell is in its top-left cell, and spans all merged columns.
      if (merge && row !== merge.y) {
        continue;
      }

      const value = merge ? this.#worksheet.getValue(merge.x, row) : values[row]![0];
      const format = (merge ? this.#worksheet.getFormattings(merge.x, row, 1, 1)[0]![0] : formattings[row]![0])!.format;
      let otherColumnsWidth = 0;

      if (merge) {
        for (let index = merge.x; index < merge.x + merge.width; index++) {
          index !== column && (otherColumnsWidth += this.#getColumnWidth(index));
        }
      }

      const { text: formatted } = formatDisplayValue(value, format);
      // Includes decorations added by CSS: currency sign (with a gap) and parentheses, or percent sign.
      const text =
        typeof value === 'number' && format === 'currency'
          ? `$ (${formatted})`
          : typeof value === 'number' && format === 'percent'
            ? `${formatted}%`
            : formatted;

      if (text) {
        for (const line of text.split('\n')) {
          width = Math.max(width, context.measureText(line).width - otherColumnsWidth);
        }
      }
    }

    return Math.ceil(width + extra + 2);
  }

  #measureContext: CanvasRenderingContext2D | null | undefined;

  /** Updates positions which depend on column widths. */
  #updateColumnLayout() {
    this.#content.style.width = `${this.#rowHeaderWidth + this.#columnOffset(this.#columnCount)}px`;

    for (let column = 0; column < this.#lockedColumnCount; column++) {
      const left = `${this.#rowHeaderWidth + this.#columnOffset(column)}px`;

      this.#columnHeaders[column]!.style.setProperty('--sticky-left', left);

      for (const { cells } of [...this.#pool, ...this.#frozenRows]) {
        cells[column]!.style.setProperty('--sticky-left', left);
      }
    }

    this.#editMode && this.#positionEditor();
    this.#updateSelection();
  }

  /**
   * Marks a cell or column header in a locked column. While locked, it sticks to the left so the column is frozen.
   */
  #initLockableColumn(element: HTMLElement, column: number) {
    const lockable = column < this.#lockedColumnCount;

    element.classList.toggle('grid__lockable', lockable);

    if (lockable) {
      // Only applied while locked, otherwise it would offset relatively positioned cells (e.g. with borders).
      element.style.setProperty('--sticky-left', `${this.#rowHeaderWidth + this.#columnOffset(column)}px`);
    } else {
      element.style.removeProperty('--sticky-left');
    }
  }

  #renderRowContent(slot: RowSlot) {
    const { cells, row } = slot;
    const worksheet = this.#worksheet;
    const merges = worksheet.getMerges(0, row, this.#columnCount, 1).sort((a, b) => a.x - b.x);
    // Frozen rows are always rendered, other rows only from the first rendered row.
    const firstRenderedRow = this.#frozenRows.includes(slot) ? 0 : this.#firstRow;
    let mergeIndex = 0;

    for (let column = 0; column < cells.length; column++) {
      const cell = cells[column]!;
      const merge = (merges[mergeIndex]?.x ?? Infinity) <= column ? merges[mergeIndex] : undefined;
      const lastColumn = merge ? merge.x + merge.width - 1 : column;
      const lastRow = merge ? merge.y + merge.height - 1 : row;
      // Content and formatting of a merged cell are in its top-left cell.
      const x = merge ? merge.x : column;
      const y = merge ? merge.y : row;

      merge && column === lastColumn && mergeIndex++;

      // A merged cell is shown as its first cell spanning all merged columns, other cells are hidden. A tall merged cell
      // is drawn by its first rendered row spanning all merged rows (also above if not rendered), other rows only keep
      // the room for it.
      const hidden = !!merge && column !== merge.x;
      const covered = !!merge && row !== Math.max(merge.y, firstRenderedRow);
      const width = `${merge ? this.#columnOffset(lastColumn + 1) - this.#columnOffset(column) : this.#getColumnWidth(column)}px`;
      const height = merge && merge.height > 1 && !covered ? `${merge.height * this.#rowHeight}px` : '';
      const top = merge && row > merge.y && !covered ? `${(merge.y - row) * this.#rowHeight}px` : '';

      const refClassName = `grid__cell-ref--${getColumnName(column).toLowerCase()}${row + 1}`;

      if (!cell.classList.contains(refClassName)) {
        [...cell.classList].forEach(
          className => className.startsWith('grid__cell-ref--') && cell.classList.remove(className)
        );
        cell.classList.add(refClassName);
      }

      if (cell.style.display !== (hidden ? 'none' : '')) {
        cell.style.display = hidden ? 'none' : '';
      }

      if (cell.style.visibility !== (covered ? 'hidden' : '')) {
        cell.style.visibility = covered ? 'hidden' : '';
      }

      if (cell.style.width !== width) {
        cell.style.width = width;
      }

      if (cell.style.height !== height) {
        cell.style.height = height;
      }

      if (cell.style.top !== top) {
        cell.style.top = top;
        cell.style.position = top ? 'relative' : '';
      }

      const value = worksheet.getValue(x, y);
      const type = typeof value;
      const formatting = worksheet.getFormattings(x, y, 1, 1)[0]![0];
      // Number formats only apply to numbers, decorations (e.g. currency sign) are added by CSS.
      const format = (type === 'number' && formatting?.format) || '';
      const { negative, text } = formatDisplayValue(value, formatting?.format);

      // Text is in an inner element, so the highlight marker only covers the text, not the whole cell.
      const content = cell.firstElementChild as HTMLElement;

      if (content.textContent !== text) {
        content.textContent = text;
      }

      if (cell.dataset['type'] !== type) {
        cell.dataset['type'] = type;
      }

      if ((cell.dataset['format'] || '') !== format) {
        format ? (cell.dataset['format'] = format) : delete cell.dataset['format'];
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

      if ((cell.dataset['align'] || '') !== align) {
        align ? (cell.dataset['align'] = align) : delete cell.dataset['align'];
      }

      const highlight = parseHighlight(formatting?.highlight);
      const background = highlight?.background || '';
      const color = highlight?.color || '';

      // Compare with the last set value, as the browser normalizes colors.
      if ((cell.dataset['background'] || '') !== background) {
        background
          ? content.style.setProperty('--marker-color', background)
          : content.style.removeProperty('--marker-color');
        content.classList.toggle('grid__cell-text--marker', !!background);
        background ? (cell.dataset['background'] = background) : delete cell.dataset['background'];
      }

      if ((cell.dataset['color'] || '') !== color) {
        cell.style.color = color;
        color ? (cell.dataset['color'] = color) : delete cell.dataset['color'];
      }

      const shadow = this.#getBorderShadow(worksheet, x, y, lastColumn, lastRow);

      // Compare with the last set value, as the browser normalizes `style.boxShadow`.
      if ((cell.dataset['border'] || '') !== shadow) {
        cell.dataset['border'] = shadow;
        cell.style.boxShadow = shadow;
      }

      const corners = this.#getBorderCorners(worksheet, x, y, lastColumn, lastRow);

      if ((cell.dataset['borderCorners'] || '') !== corners) {
        cell.dataset['borderCorners'] = corners;
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
   * `lastColumn` and `lastRow` are the last column and row of a merged cell.
   */
  #getBorderShadow(
    worksheet: Worksheet,
    column: number,
    row: number,
    lastColumn: number = column,
    lastRow: number = row
  ) {
    const inside = BORDER_WIDTH - BORDER_OUTSIDE_WIDTH;
    const shadows: string[] = [];

    // Edges along a merged cell only span its full size if present in every column or row, otherwise drawn by segments.
    this.#every(row, lastRow, r => hasVerticalEdge(worksheet, lastColumn, r)) &&
      shadows.push(`inset -${inside}px 0 ${BORDER_COLOR}`);
    this.#every(column, lastColumn, c => hasHorizontalEdge(worksheet, c, lastRow)) &&
      shadows.push(`inset 0 -${inside}px ${BORDER_COLOR}`);
    this.#every(row, lastRow, r => hasVerticalEdge(worksheet, column - 1, r)) &&
      shadows.push(`inset ${BORDER_OUTSIDE_WIDTH}px 0 ${BORDER_COLOR}`);
    this.#every(column, lastColumn, c => hasHorizontalEdge(worksheet, c, row - 1)) &&
      shadows.push(`inset 0 ${BORDER_OUTSIDE_WIDTH}px ${BORDER_COLOR}`);

    return shadows.length ? [...shadows, 'var(--grid-line-shadow)'].join(', ') : '';
  }

  /**
   * Returns the background layers filling the parts of border junctions (where grid lines cross) inside a cell, so
   * borders meeting at a corner join without a notch. Empty if none.
   *
   * `lastColumn` and `lastRow` are the last column and row of a merged cell.
   */
  #getBorderCorners(
    worksheet: Worksheet,
    column: number,
    row: number,
    lastColumn: number = column,
    lastRow: number = row
  ) {
    const inside = BORDER_WIDTH - BORDER_OUTSIDE_WIDTH;
    const outside = BORDER_OUTSIDE_WIDTH;
    const fill = `linear-gradient(${BORDER_COLOR}, ${BORDER_COLOR})`;
    const layers: string[] = [];

    hasJunction(worksheet, lastColumn, lastRow) &&
      layers.push(`${fill} right bottom / ${inside}px ${inside}px no-repeat`);
    hasJunction(worksheet, column - 1, lastRow) &&
      layers.push(`${fill} left bottom / ${outside}px ${inside}px no-repeat`);
    hasJunction(worksheet, lastColumn, row - 1) &&
      layers.push(`${fill} right top / ${inside}px ${outside}px no-repeat`);
    hasJunction(worksheet, column - 1, row - 1) &&
      layers.push(`${fill} left top / ${outside}px ${outside}px no-repeat`);

    // Segments of edges present along only some columns or rows of a merged cell.
    const bottomPartial = !this.#every(column, lastColumn, c => hasHorizontalEdge(worksheet, c, lastRow));
    const topPartial = !this.#every(column, lastColumn, c => hasHorizontalEdge(worksheet, c, row - 1));
    const rightPartial = !this.#every(row, lastRow, r => hasVerticalEdge(worksheet, lastColumn, r));
    const leftPartial = !this.#every(row, lastRow, r => hasVerticalEdge(worksheet, column - 1, r));

    for (let c = column; c <= lastColumn; c++) {
      const x = this.#columnOffset(c) - this.#columnOffset(column);
      const width = this.#getColumnWidth(c);

      bottomPartial &&
        hasHorizontalEdge(worksheet, c, lastRow) &&
        layers.push(`${fill} ${x}px bottom / ${width}px ${inside}px no-repeat`);
      topPartial &&
        hasHorizontalEdge(worksheet, c, row - 1) &&
        layers.push(`${fill} ${x}px top / ${width}px ${outside}px no-repeat`);
    }

    for (let r = row; r <= lastRow; r++) {
      const y = (r - row) * this.#rowHeight;

      rightPartial &&
        hasVerticalEdge(worksheet, lastColumn, r) &&
        layers.push(`${fill} right ${y}px / ${inside}px ${this.#rowHeight}px no-repeat`);
      leftPartial &&
        hasVerticalEdge(worksheet, column - 1, r) &&
        layers.push(`${fill} left ${y}px / ${outside}px ${this.#rowHeight}px no-repeat`);
    }

    return layers.join(', ');
  }

  /**
   * Returns `true` if the predicate is true for every column or row from `first` to `last`.
   */
  #every(first: number, last: number, predicate: (index: number) => boolean) {
    for (let index = first; index <= last; index++) {
      if (!predicate(index)) {
        return false;
      }
    }

    return true;
  }

  #renderRowSelection({ header, row }: Pick<RowSlot, 'header' | 'row'>) {
    const { top, bottom } = this.#worksheet.selection.range;

    header.classList.toggle('grid__row-header--selected', row >= top && row <= bottom);
  }

  #updateSelection() {
    const rowHeight = this.#rowHeight;
    const rowHeaderWidth = this.#rowHeaderWidth;
    const { left, top, right, bottom } = this.#worksheet.selection.range;
    const rangeStyle = this.#rangeOverlay.style;
    const activeStyle = this.#activeOverlay.style;
    const anchorEnd = this.#toMergeEnd(this.#worksheet.selection.activeCell);

    rangeStyle.left = `${rowHeaderWidth + this.#columnOffset(left)}px`;
    rangeStyle.top = `${top * rowHeight}px`;
    rangeStyle.width = `${this.#columnOffset(right + 1) - this.#columnOffset(left)}px`;
    rangeStyle.height = `${(bottom - top + 1) * rowHeight}px`;
    this.#rangeOverlay.hidden =
      left === this.#worksheet.selection.activeCell.column &&
      right === anchorEnd.column &&
      top === this.#worksheet.selection.activeCell.row &&
      bottom === anchorEnd.row;

    activeStyle.left = `${rowHeaderWidth + this.#columnOffset(this.#worksheet.selection.activeCell.column)}px`;
    activeStyle.width = `${this.#columnOffset(anchorEnd.column + 1) - this.#columnOffset(this.#worksheet.selection.activeCell.column)}px`;
    activeStyle.top = `${this.#worksheet.selection.activeCell.row * rowHeight}px`;
    activeStyle.height = `${(anchorEnd.row - this.#worksheet.selection.activeCell.row + 1) * rowHeight}px`;

    this.#columnHeaders.forEach((header, column) =>
      header.classList.toggle('grid__column-header--selected', column >= left && column <= right)
    );

    this.#pool.forEach(slot => slot.row !== -1 && this.#renderRowSelection(slot));
    this.#frozenRows.forEach(slot => this.#renderRowSelection(slot));
  }

  // #endregion

  // #region Keyboard

  #clamp({ column, row }: CellPosition): CellPosition {
    return {
      column: Math.max(0, Math.min(this.#columnCount - 1, column)),
      row: Math.max(0, Math.min(this.#rowCount - 1, row))
    };
  }

  /**
   * Moves the active cell, or extends the selection.
   */
  #move(deltaColumn: number, deltaRow: number, extend: boolean) {
    if (extend) {
      return this.#select(this.#worksheet.selection.activeCell, this.#extendFocus(deltaColumn, deltaRow));
    }

    const from = this.#worksheet.selection.activeCell;
    const preferred = this.#preferred;
    const end = this.#toMergeEnd(from);
    // Moving right or down skips the rest of a merged cell. Moving vertically out of a merged cell keeps the column
    // entered from, moving horizontally keeps the row entered from.
    const column =
      deltaColumn > 0
        ? end.column
        : deltaRow && preferred.column >= from.column && preferred.column <= end.column
          ? preferred.column
          : from.column;
    const row =
      deltaRow > 0
        ? end.row
        : deltaColumn && preferred.row >= from.row && preferred.row <= end.row
          ? preferred.row
          : from.row;

    this.#select(this.#clamp({ column: column + deltaColumn, row: row + deltaRow }));
  }

  /**
   * Returns the new focus when extending the selection. The focus moves in the direction until the selection changes,
   * skipping positions that select the same range because merged cells already cover them.
   */
  #extendFocus(deltaColumn: number, deltaRow: number): CellPosition {
    const selection = this.#worksheet.selection.range;
    let focus = this.#focus;

    for (;;) {
      const next = { column: focus.column + deltaColumn, row: focus.row + deltaRow };

      if (next.column < 0 || next.column >= this.#columnCount || next.row < 0 || next.row >= this.#rowCount) {
        return this.#focus;
      }

      const range = this.#rangeOf(this.#worksheet.selection.activeCell, next);

      if (
        range.left !== selection.left ||
        range.right !== selection.right ||
        range.top !== selection.top ||
        range.bottom !== selection.bottom
      ) {
        return next;
      }

      focus = next;
    }
  }

  /**
   * Moves the active cell within the selection (Enter / Tab), wrapping around like Excel when a range is selected.
   */
  #moveWithinSelection(deltaColumn: number, deltaRow: number) {
    const { left, top, right, bottom } = this.#worksheet.selection.range;

    const end = this.#toMergeEnd(this.#worksheet.selection.activeCell);

    // A single (possibly merged) cell is selected.
    if (
      left === this.#worksheet.selection.activeCell.column &&
      right === end.column &&
      top === this.#worksheet.selection.activeCell.row &&
      bottom === end.row
    ) {
      return this.#move(deltaColumn, deltaRow, false);
    }

    let { column, row } = this.#worksheet.selection.activeCell;

    // Skips cells covered by a merged cell other than its top-left cell, which is visited once.
    for (let count = (right - left + 1) * (bottom - top + 1); count > 0; count--) {
      const end = this.#toMergeEnd({ column, row });

      // Moving right or down skips the rest of a merged cell.
      column = (deltaColumn > 0 ? end.column : column) + deltaColumn;
      row = (deltaRow > 0 ? end.row : row) + deltaRow;

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

      const start = this.#toMergeStart({ column, row });

      if (start.column === column && start.row === row) {
        break;
      }
    }

    // Keep the selected range, only move the active cell within it.
    this.#editMode && this.#commitEdit();
    this.#setSelectionObject(new WorksheetSelection(this.#toMergeStart({ column, row }), { left, top, right, bottom }));
    this.#scrollIntoView(this.#worksheet.selection.activeCell);
  }

  #handleKeyDown(event: KeyboardEvent) {
    if (event.target !== this || event.isComposing) {
      return;
    }

    const { key, shiftKey } = event;
    const ctrlKey = event.ctrlKey || event.metaKey;
    const frozenHeight = this.#locked ? this.#lockedRowCount * this.#rowHeight : 0;
    const pageRows = Math.max(
      1,
      Math.floor((this.clientHeight - this.#headerHeight - frozenHeight) / this.#rowHeight) - 1
    );
    const { column, row } = shiftKey ? this.#focus : this.#worksheet.selection.activeCell;

    // Alt+Arrow moves the active cell within the selection without changing the selected range.
    if (event.altKey && !shiftKey && !ctrlKey && key.startsWith('Arrow')) {
      event.preventDefault();

      const { left, top, right, bottom } = this.#worksheet.selection.range;
      const end = this.#toMergeEnd(this.#worksheet.selection.activeCell);
      const next = {
        column:
          key === 'ArrowLeft'
            ? Math.max(left, column - 1)
            : key === 'ArrowRight'
              ? Math.min(right, end.column + 1)
              : column,
        row: key === 'ArrowUp' ? Math.max(top, row - 1) : key === 'ArrowDown' ? Math.min(bottom, end.row + 1) : row
      };
      const activeCell = this.#toMergeStart(next);

      this.#setSelectionObject(new WorksheetSelection(activeCell, { left, top, right, bottom }));
      this.#scrollIntoView(activeCell);

      return;
    }

    let delta: [number, number] | undefined;

    switch (key) {
      case 'ArrowUp':
        delta = [0, ctrlKey ? -row : -1];
        break;

      case 'ArrowDown':
        delta = [0, ctrlKey ? this.#rowCount - 1 - row : 1];
        break;

      case 'ArrowLeft':
        delta = [ctrlKey ? -column : -1, 0];
        break;

      case 'ArrowRight':
        delta = [ctrlKey ? this.#columnCount - 1 - column : 1, 0];
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
        delta = [this.#columnCount - 1 - column, ctrlKey ? this.#rowCount - 1 - row : 0];
        break;

      case 'Enter':
        event.preventDefault();
        this.#moveWithinSelection(0, shiftKey ? -1 : 1);

        return;

      case 'Tab':
        event.preventDefault();
        this.#moveWithinSelection(shiftKey ? -1 : 1, 0);

        return;

      case 'Escape':
        event.preventDefault();
        this.#select(this.#worksheet.selection.activeCell);

        return;

      case 'F2':
        event.preventDefault();
        this.#startEdit();

        return;

      case 'Delete':
      case 'Backspace':
        event.preventDefault();
        this.#clearSelection();

        return;

      default:
        if (ctrlKey && !event.altKey && key.toLowerCase() === 'a') {
          event.preventDefault();
          this.#selectKeepingScroll({ column: 0, row: 0 }, { column: this.#columnCount - 1, row: this.#rowCount - 1 });
        } else if (key.length === 1 && !ctrlKey && !event.altKey) {
          // Printable character: start editing, replacing the content (same as Excel).
          event.preventDefault();
          this.#startEdit(key);
        }

        return;
    }

    event.preventDefault();
    this.#move(delta[0], delta[1], shiftKey);
  }

  // #endregion

  // #region Pointer

  #cellFromPoint({ clientX, clientY }: PointerEvent | MouseEvent): CellPosition {
    const element = this;
    const headerHeight = this.#headerHeight;
    const rowHeaderWidth = this.#rowHeaderWidth;
    const rowHeight = this.#rowHeight;
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

  #handlePointerDown(event: PointerEvent) {
    const target = event.target as HTMLElement;

    if (event.button !== 0 || target === this.#editor) {
      return;
    }

    const scrollerRect = this.getBoundingClientRect();

    // Ignore clicks on the native scrollbars.
    if (
      event.clientX - scrollerRect.left >= this.clientLeft + this.clientWidth ||
      event.clientY - scrollerRect.top >= this.clientTop + this.clientHeight
    ) {
      return;
    }

    event.preventDefault();
    this.focus({ preventScroll: true });

    const resizeColumn = this.#resizeColumnFromPoint(event);

    if (resizeColumn >= 0) {
      this.#dragging = 'resize';
      this.#resizing = {
        columns: this.#columnsToResize(resizeColumn),
        startX: event.clientX,
        startWidth: this.#getColumnWidth(resizeColumn)
      };
      this.setPointerCapture(event.pointerId);

      return;
    }

    const cell = this.#cellFromPoint(event);
    const { shiftKey } = event;

    if (target.closest('.grid__corner')) {
      this.#selectKeepingScroll({ column: 0, row: 0 }, { column: this.#columnCount - 1, row: this.#rowCount - 1 });

      return;
    } else if (target.closest('.grid__column-header')) {
      const column = Number((target.closest('.grid__column-header') as HTMLElement).dataset['column']);

      this.#dragging = 'column';
      shiftKey
        ? this.#selectKeepingScroll(this.#worksheet.selection.activeCell, { column, row: this.#rowCount - 1 })
        : this.#selectKeepingScroll({ column, row: 0 }, { column, row: this.#rowCount - 1 });
    } else if (target.closest('.grid__row-header')) {
      this.#dragging = 'row';
      shiftKey
        ? this.#selectKeepingScroll(this.#worksheet.selection.activeCell, {
            column: this.#columnCount - 1,
            row: cell.row
          })
        : this.#selectKeepingScroll({ column: 0, row: cell.row }, { column: this.#columnCount - 1, row: cell.row });
    } else {
      this.#dragging = 'cell';
      shiftKey ? this.#select(this.#worksheet.selection.activeCell, cell) : this.#select(cell);
    }

    this.setPointerCapture(event.pointerId);
  }

  /**
   * Selects without scrolling to the focus cell, used for whole row/column selection.
   */
  #selectKeepingScroll(anchor: CellPosition, focus: CellPosition) {
    this.#editMode && this.#commitEdit();
    this.#setSelection(anchor, focus);
  }

  #handlePointerMove(event: PointerEvent) {
    if (!this.#dragging) {
      this.classList.toggle('grid--resize-hover', this.#resizeColumnFromPoint(event) >= 0);

      return;
    }

    if (this.#dragging === 'resize') {
      const { columns, startX, startWidth } = this.#resizing!;

      columns.forEach(column => this.#setColumnWidth(column, startWidth + event.clientX - startX));

      return;
    }

    const cell = this.#cellFromPoint(event);

    if (this.#dragging === 'column') {
      this.#selectKeepingScroll(this.#worksheet.selection.activeCell, { column: cell.column, row: this.#rowCount - 1 });
    } else if (this.#dragging === 'row') {
      this.#selectKeepingScroll(this.#worksheet.selection.activeCell, { column: this.#columnCount - 1, row: cell.row });
    } else if (cell.column !== this.#focus.column || cell.row !== this.#focus.row) {
      this.#select(this.#worksheet.selection.activeCell, cell);
    }
  }

  #handlePointerUp(event: PointerEvent) {
    if (this.#dragging) {
      this.#dragging = '';
      this.#resizing = undefined;
      this.hasPointerCapture(event.pointerId) && this.releasePointerCapture(event.pointerId);
    }
  }

  /**
   * Returns the column whose right border in the column header is under the pointer, or -1. Always -1 while locked.
   */
  #resizeColumnFromPoint({ clientX, clientY }: PointerEvent | MouseEvent) {
    if (this.#locked) {
      return -1;
    }

    const element = this;
    const rowHeaderWidth = this.#rowHeaderWidth;
    const scrollerRect = element.getBoundingClientRect();
    const y = clientY - scrollerRect.top - element.clientTop;

    if (y < 0 || y >= this.#headerHeight) {
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

    for (let column = 0; column < this.#columnCount; column++) {
      offset += this.#getColumnWidth(column);

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
   */
  #columnsToResize(column: number): number[] {
    const { left, top, right, bottom } = this.#worksheet.selection.range;

    if (top === 0 && bottom === this.#rowCount - 1 && column >= left && column <= right) {
      return Array.from({ length: right - left + 1 }, (_, index) => left + index);
    }

    return [column];
  }

  #handleDoubleClick(event: MouseEvent) {
    const target = event.target as HTMLElement;
    const resizeColumn = this.#resizeColumnFromPoint(event);

    if (resizeColumn >= 0) {
      this.#autoFitColumns(this.#columnsToResize(resizeColumn));

      return;
    }

    // With pointer capture, the target may be the grid itself, so hit-test using coordinates.
    if (
      target !== this.#editor &&
      !target.closest('.grid__header, .grid__row-header') &&
      event.clientY >= this.#body.getBoundingClientRect().top &&
      event.clientX >= this.getBoundingClientRect().left + this.#rowHeaderWidth
    ) {
      this.#startEdit();
    }
  }

  // #endregion

  // #region Editing

  /**
   * Opens the in-cell editor on the active cell.
   *
   * When `initialText` is specified, replaces the cell content ("enter" mode). Otherwise, edits the current content ("edit" mode).
   * Returns `false` if the cell is locked.
   */
  #startEdit(initialText?: string): boolean {
    const { column, row } = this.#worksheet.selection.activeCell;

    if (!this.#isEditable(column, row)) {
      this.dispatchEvent(new CustomEvent('lockedit', { detail: { column, row } }));

      return false;
    }

    this.#scrollIntoView(this.#worksheet.selection.activeCell);

    const editor = this.#editor;
    const value = this.#worksheet.getValue(column, row);

    this.#positionEditor();
    editor.value = initialText ?? formatValue(value);
    editor.hidden = false;
    editor.focus({ preventScroll: true });
    editor.setSelectionRange(editor.value.length, editor.value.length);

    this.#editCell = { column, row };
    this.#editMode = initialText === undefined ? 'edit' : 'enter';

    return true;
  }

  /** Commits the editor value to the active cell. */
  #positionEditor() {
    const { column, row } = this.#worksheet.selection.activeCell;
    const style = this.#editor.style;

    style.left = `${this.#rowHeaderWidth + this.#columnOffset(column)}px`;
    style.top = `${row * this.#rowHeight}px`;
    const end = this.#toMergeEnd(this.#worksheet.selection.activeCell);

    style.width = `${this.#columnOffset(end.column + 1) - this.#columnOffset(column)}px`;
    style.height = `${(end.row - row + 1) * this.#rowHeight}px`;
  }

  #commitEdit() {
    if (!this.#editMode) {
      return;
    }

    const { column, row } = this.#editCell;
    const text = this.#editor.value;

    this.#closeEditor();

    if (this.#isEditable(column, row)) {
      // Leading apostrophe forces the value to be text, same as Excel.
      // Leading "=" evaluates a simple arithmetic formula and stores only the result; invalid formulas are kept as text.
      const result = text.startsWith('=') ? calculate(text.slice(1)) : undefined;

      this.#worksheet.edit(edit =>
        edit.setValue(column, row, text.startsWith("'") ? text.slice(1) : (result ?? inferValue(text)))
      );
    }
  }

  /** Closes the editor without saving. */
  #cancelEdit() {
    this.#editMode && this.#closeEditor();
  }

  #closeEditor() {
    const hadFocus = document.activeElement === this.#editor;

    this.#editMode = '';
    this.#editor.hidden = true;

    hadFocus && this.focus({ preventScroll: true });
  }

  #handleEditorKeyDown(event: KeyboardEvent) {
    if (event.isComposing) {
      return;
    }

    const { key, shiftKey } = event;

    if (key === 'Escape') {
      event.preventDefault();
      this.#cancelEdit();
    } else if (key === 'Enter' && !event.altKey) {
      event.preventDefault();
      this.#commitEdit();
      this.#moveWithinSelection(0, shiftKey ? -1 : 1);
    } else if (key === 'Tab') {
      event.preventDefault();
      this.#commitEdit();
      this.#moveWithinSelection(shiftKey ? -1 : 1, 0);
    } else if (key === 'F2') {
      event.preventDefault();
      this.#editMode = this.#editMode === 'edit' ? 'enter' : 'edit';
    } else if (this.#editMode === 'enter' && key.startsWith('Arrow')) {
      event.preventDefault();
      this.#commitEdit();
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
  #clearSelection() {
    const worksheet = this.#worksheet;

    const { left, top, right, bottom } = this.#worksheet.selection.range;
    let skipped = false;

    worksheet.edit(edit => {
      for (let row = top; row <= bottom; row++) {
        for (let column = left; column <= right; column++) {
          if (this.#isEditable(column, row)) {
            edit.setValue(column, row, '');
          } else {
            skipped = true;
          }
        }
      }
    });

    skipped && this.dispatchEvent(new CustomEvent('lockedit', { detail: this.#worksheet.selection.activeCell }));
  }

  // #endregion

  // #region Clipboard

  #getSelectedValues(): CellValue[][] {
    const worksheet = this.#worksheet;
    const { left, top, right, bottom } = this.#worksheet.selection.range;

    const values: CellValue[][] = [];

    if (worksheet) {
      for (let row = top; row <= bottom; row++) {
        const rowValues: CellValue[] = [];

        for (let column = left; column <= right; column++) {
          rowValues.push(worksheet.getValue(column, row) ?? '');
        }

        values.push(rowValues);
      }
    }

    return values;
  }

  #handleCopy(event: ClipboardEvent, cut: boolean) {
    if (event.target !== this || !event.clipboardData) {
      return;
    }

    event.preventDefault();

    const values = this.#getSelectedValues();
    const { left, top, right, bottom } = this.#worksheet.selection.range;
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

    cut && this.#clearSelection();
  }

  #handlePaste(event: ClipboardEvent) {
    const worksheet = this.#worksheet;

    if (event.target !== this || !event.clipboardData) {
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
    const { left, top, right, bottom } = this.#worksheet.selection.range;
    const selectedRows = bottom - top + 1;
    const selectedColumns = right - left + 1;

    // Same as Excel: if the selection is a multiple of the copied data, the data is repeated to fill the selection.
    const tile = selectedRows % dataRows === 0 && selectedColumns % dataColumns === 0;
    const pasteRows = Math.min(tile ? selectedRows : dataRows, this.#rowCount - top);
    const pasteColumns = Math.min(tile ? selectedColumns : dataColumns, this.#columnCount - left);
    let skipped = false;

    worksheet.edit(edit => {
      for (let rowOffset = 0; rowOffset < pasteRows; rowOffset++) {
        const rowValues = values[rowOffset % dataRows]!;

        for (let columnOffset = 0; columnOffset < pasteColumns; columnOffset++) {
          const column = left + columnOffset;
          const row = top + rowOffset;

          if (this.#isEditable(column, row)) {
            edit.setValue(column, row, rowValues[columnOffset % dataColumns] ?? '');

            if (formattings) {
              const formattingRow = formattings[rowOffset % dataRows];

              edit.setFormattings([[formattingRow?.[columnOffset % (formattingRow?.length || 1)]]], column, row);
            }
          } else {
            skipped = true;
          }
        }
      }
    });

    const anchor = { column: left, row: top };

    this.#selectKeepingScroll(anchor, { column: left + pasteColumns - 1, row: top + pasteRows - 1 });

    skipped && this.dispatchEvent(new CustomEvent('lockedit', { detail: anchor }));
  }

  // #endregion
}

customElements.get('worksheet-app') || customElements.define('worksheet-app', WorksheetApp);

export { WorksheetApp };
