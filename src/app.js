import { formatValue } from './csv.js';
import { HIGHLIGHTS } from './highlights.js';
import { cellsToPNG } from './printCells.js';
import { MAX_WORKSHEETS } from './workbook/constants.js';
import { getColumnName } from './worksheet/getColumnName.js';
import { Workbook } from './workbook/Workbook.js';
import { WorksheetElement } from './worksheet/WorksheetElement.js';
import { WorksheetSelection } from './worksheet/WorksheetSelection.js';
import { mountUI } from './ui/mountUI.jsx';

/** @typedef {import('./worksheet/Worksheet.js').Worksheet} Worksheet */

/**
 * Number of leading rows and columns that become read-only, and frozen on screen, when the lock button is pressed.
 * Change these to protect a different range, e.g. 2 header rows and 1 ID column.
 */
const LOCKED_ROW_COUNT = 1;
const LOCKED_COLUMN_COUNT = 1;

/** @typedef {import('./ui/Toolbar.jsx').ToolbarActions} ToolbarActions */

const fileInput = /** @type {HTMLInputElement} */ (document.getElementById('file-input'));
/** @type {Record<import('./ui/Toolbar.jsx').CellStyle, boolean>} */
const NO_STYLES = { bold: false, italic: false, underline: false, strikethrough: false };
/** @type {readonly import('./worksheet/Worksheet.js').CellAlign[]} */
const ALIGNS = ['left', 'center', 'right'];
/** @type {readonly import('./worksheet/Worksheet.js').CellNumberFormat[]} */
const FORMATS = ['currency', 'percent', 'date', 'datetime', 'duration'];

/** @type {ToolbarActions} */
const toolbarActions = {
  merge: () => mergeSelection(),
  open: () => fileInput.click(),
  print: () => printSelection(),
  redo: () => undoOrRedo('redo'),
  saveCSV: () => saveCSV(),
  saveZip: () => saveZip(),
  setAlign: align => formatSelection(formatting => ({ ...formatting, align })),
  setBorder: preset => setBorder(preset),
  setFormat: format => formatSelection(formatting => ({ ...formatting, format })),
  setHighlight: highlight => formatSelection(formatting => ({ ...formatting, highlight })),
  setLocked: nextLocked => setLocked(nextLocked),
  toggleStyle: style => {
    const value = !isFormattingAppliedToSelection(formatting => !!formatting?.[style]);

    formatSelection(formatting => ({ ...formatting, [style]: value }));
  },
  undo: () => undoOrRedo('undo'),
  unformat: () => formatSelection(() => null)
};

const ui = mountUI(
  /** @type {HTMLElement} */ (document.getElementById('root')),
  toolbarActions,
  {
    align: undefined,
    canRedo: false,
    canUndo: false,
    format: undefined,
    highlight: undefined,
    locked: false,
    merged: false,
    nameBox: 'A1',
    status: '',
    styles: NO_STYLES,
    valueBox: ''
  },
  // Menus return focus to the grid after closing.
  {
    get current() {
      return activeElement ?? null;
    }
  }
);

const gridContainer = /** @type {HTMLElement} */ (document.getElementById('grid-container'));
const sheetTabs = /** @type {HTMLElement} */ (document.getElementById('sheet-tabs'));
const addSheetButton = /** @type {HTMLButtonElement} */ (document.getElementById('add-sheet-button'));

/** `<compulim-worksheet>` of the active worksheet. @type {WorksheetElement | undefined} */
let activeElement;

let locked = false;

/** @param {Worksheet} worksheet */
function getElement(worksheet) {
  return /** @type {WorksheetElement} */ (workbook.getWorksheetElement(worksheet));
}

/**
 * Shows the `<compulim-worksheet>` instantiated by the workbook for a worksheet.
 *
 * @param {Worksheet} worksheet
 */
function mountWorksheet(worksheet) {
  const element = getElement(worksheet);

  element.setAttribute('locked-column-count', String(LOCKED_COLUMN_COUNT));
  element.setAttribute('locked-row-count', String(LOCKED_ROW_COUNT));
  element.toggleAttribute('locked', locked);
  element.hidden = true;
  element.addEventListener('selectionchange', () => element === activeElement && updateSelectionInfo());
  element.addEventListener('lockedit', () => showStatus('Some cells are locked and cannot be edited.'));
  element.addEventListener('historychange', () => element === activeElement && updateUndoButtons());
  gridContainer.append(element);
}

/** @param {Event} event */
function handleWorksheetAdd(event) {
  mountWorksheet(/** @type {import('./worksheet/WorksheetEvent.js').WorksheetEvent} */ (event).worksheet);
}

/** Active cell and selection of the active worksheet. */
function getSelection() {
  return activeElement ? activeElement.selection : new WorksheetSelection({ column: 0, row: 0 });
}

/**
 * Returns `true` if the cell can be edited by the user (i.e. not locked).
 *
 * @param {number} column
 * @param {number} row
 */
function isEditable(column, row) {
  return !(locked && (column < LOCKED_COLUMN_COUNT || row < LOCKED_ROW_COUNT));
}

/** @type {Workbook} */
let workbook;

/** @type {ReturnType<typeof setTimeout> | undefined} */
let statusTimeout;

/** @param {string} message */
function showStatus(message) {
  clearTimeout(statusTimeout);
  ui.update({ status: message });
  statusTimeout = setTimeout(() => ui.update({ status: '' }), 4000);
}

/**
 * @param {Blob} blob
 * @param {string} fileName
 */
function download(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');

  anchor.href = url;
  anchor.download = fileName;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function renderTabs() {
  const active = activeElement?.worksheet;
  const worksheets = workbook.worksheets;

  sheetTabs.replaceChildren(
    ...worksheets.map(worksheet => {
      const tab = document.createElement('div');
      const label = document.createElement('span');

      tab.className = 'sheet-tab';
      tab.setAttribute('role', 'tab');
      tab.setAttribute('aria-selected', String(worksheet === active));
      tab.title = 'Click to select, double-click to rename';
      const name = workbook.getWorksheetName(worksheet);

      label.textContent = name;
      tab.append(label);

      tab.addEventListener('click', () => activateWorksheet(worksheet));
      tab.addEventListener('dblclick', () => renameWorksheet(worksheet));

      if (worksheets.length > 1) {
        const close = document.createElement('button');

        close.type = 'button';
        close.className = 'sheet-tab__close';
        close.textContent = '×';
        close.title = `Delete "${name}"`;
        close.setAttribute('aria-label', `Delete sheet ${name}`);
        close.addEventListener('click', event => {
          event.stopPropagation();
          deleteWorksheet(worksheet);
        });
        tab.append(close);
      }

      return tab;
    })
  );

  addSheetButton.disabled = worksheets.length >= MAX_WORKSHEETS;
  sheetTabs.querySelector('[aria-selected="true"]')?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
}

/** @param {Worksheet} worksheet */
function activateWorksheet(worksheet) {
  const element = getElement(worksheet);

  for (const other of workbook.worksheets) {
    getElement(other).hidden = other !== worksheet;
  }

  activeElement = element;
  renderTabs();
  updateUndoButtons();
  updateFormattingButtons();
  activeElement?.focus();
}

/** @param {Worksheet} worksheet */
function renameWorksheet(worksheet) {
  const name = prompt('Rename sheet', workbook.getWorksheetName(worksheet));

  if (name === null) {
    return;
  }

  try {
    workbook.renameWorksheet(worksheet, name.trim());
  } catch (error) {
    alert(/** @type {Error} */ (error).message);
  }
}

/** @param {Worksheet} worksheet */
function deleteWorksheet(worksheet) {
  if (workbook.worksheets.length <= 1 || !confirm(`Delete sheet "${workbook.getWorksheetName(worksheet)}"?`)) {
    return;
  }

  const index = workbook.indexOf(worksheet);

  workbook.removeWorksheet(worksheet);

  if (activeElement?.worksheet === worksheet) {
    activateWorksheet(workbook.worksheets[Math.min(index, workbook.worksheets.length - 1)]);
  }
}

/** @param {Workbook} nextWorkbook */
function setWorkbook(nextWorkbook) {
  if (!nextWorkbook.worksheets.length) {
    nextWorkbook.addWorksheet();
  }

  if (workbook) {
    workbook.removeEventListener('worksheetadd', renderTabs);
    workbook.removeEventListener('worksheetremove', renderTabs);
    workbook.removeEventListener('worksheetadd', handleWorksheetAdd);
    workbook.removeEventListener('worksheetrename', renderTabs);
    workbook.removeEventListener('change', updateUndoButtons);
    workbook.removeEventListener('change', updateFormattingButtons);
  }

  workbook = nextWorkbook;
  workbook.addEventListener('worksheetadd', renderTabs);
  workbook.addEventListener('worksheetremove', renderTabs);
  workbook.addEventListener('worksheetadd', handleWorksheetAdd);
  workbook.addEventListener('worksheetrename', renderTabs);
  workbook.addEventListener('change', updateUndoButtons);
  workbook.addEventListener('change', updateFormattingButtons);

  // Expose for scripting from the browser console, e.g. `app.workbook.worksheets[0].edit(edit => edit.setValue(0, 0, 'Hello'))`.
  Object.assign(globalThis, { app: { workbook } });

  gridContainer.replaceChildren();
  workbook.worksheets.forEach(mountWorksheet);
  activateWorksheet(workbook.worksheets[0]);
}

function updateSelectionInfo() {
  const { left, top, right, bottom } = getSelection().range;
  const { column, row } = getSelection().activeCell;
  const topLeft = `${getColumnName(left)}${top + 1}`;

  ui.update({
    nameBox: left === right && top === bottom ? topLeft : `${topLeft}:${getColumnName(right)}${bottom + 1}`,
    valueBox: activeElement ? formatValue(activeElement.worksheet.getValue(column, row)) : ''
  });
  updateFormattingButtons();
}

/**
 * Whether the predicate holds for the formatting of every contentful cell in the selection.
 * If the selection has no contentful cells, checks the active cell instead.
 *
 * @param {(formatting: import('./worksheet/Worksheet.js').CellFormatting | undefined) => boolean} predicate
 */
function isFormattingAppliedToSelection(predicate) {
  const worksheet = activeElement?.worksheet;

  if (!worksheet) {
    return false;
  }

  const { activeCell, range } = getSelection();
  const width = range.right - range.left + 1;
  const height = range.bottom - range.top + 1;
  const formattings = worksheet.getFormattings(range.left, range.top, width, height);
  let hasContent = false;

  for (let row = range.top; row <= range.bottom; row++) {
    for (let column = range.left; column <= range.right; column++) {
      if ((worksheet.getValue(column, row) ?? '') !== '') {
        hasContent = true;

        if (!predicate(formattings[row - range.top]?.[column - range.left])) {
          return false;
        }
      }
    }
  }

  return hasContent || predicate(worksheet.getFormattings(activeCell.column, activeCell.row, 1, 1)[0][0]);
}

function updateFormattingButtons() {
  /** @type {Record<import('./ui/Toolbar.jsx').CellStyle, boolean>} */
  const styles = { ...NO_STYLES };

  for (const style of /** @type {const} */ (['bold', 'italic', 'underline', 'strikethrough'])) {
    styles[style] = isFormattingAppliedToSelection(formatting => !!formatting?.[style]);
  }

  ui.update({
    align: ALIGNS.find(align => isFormattingAppliedToSelection(formatting => formatting?.align === align)),
    format: FORMATS.find(format => isFormattingAppliedToSelection(formatting => formatting?.format === format)),
    highlight: HIGHLIGHTS.find(({ name }) =>
      isFormattingAppliedToSelection(formatting => formatting?.highlight === name)
    )?.name,
    merged: isSelectionMerged(),
    styles
  });
}

/** Whether every selected cell is in a merged cell. */
function isSelectionMerged() {
  const worksheet = activeElement?.worksheet;

  if (!worksheet) {
    return false;
  }

  const { left, top, right, bottom } = getSelection().range;

  let covered = 0;

  for (const { x, y, width, height } of worksheet.getMerges(left, top, right - left + 1, bottom - top + 1)) {
    covered +=
      (Math.min(x + width - 1, right) - Math.max(x, left) + 1) *
      (Math.min(y + height - 1, bottom) - Math.max(y, top) + 1);
  }

  // Merged cells never overlap, so all cells are merged if they cover the whole selection.
  return covered === (right - left + 1) * (bottom - top + 1);
}

/** @param {import('./worksheet/Worksheet.js').FormattingUpdater} updater */
function formatSelection(updater) {
  const worksheet = activeElement?.worksheet;

  if (worksheet) {
    const { left, top, right, bottom } = getSelection().range;

    worksheet.edit(edit => edit.setFormattings(updater, left, top, right - left + 1, bottom - top + 1));
    updateFormattingButtons();
    updateUndoButtons();
  }

  activeElement?.focus();
}

/** @param {import('./worksheet/Worksheet.js').CellAlign} align */
function toggleAlign(align) {
  // Toggling a pushed alignment goes back to the default (type-based) alignment.
  const nextAlign = isFormattingAppliedToSelection(formatting => formatting?.align === align) ? undefined : align;

  formatSelection(formatting => ({ ...formatting, align: nextAlign }));
}

/** @type {Record<string, () => void>} */
const formattingHotkeys = {
  b: () => toolbarActions.toggleStyle('bold'),
  i: () => toolbarActions.toggleStyle('italic'),
  u: () => toolbarActions.toggleStyle('underline'),
  l: () => toggleAlign('left'),
  e: () => toggleAlign('center'),
  r: () => toggleAlign('right'),
  m: () => mergeSelection()
};

gridContainer.addEventListener('keydown', event => {
  const target = /** @type {HTMLElement} */ (event.target);

  if (
    (event.ctrlKey || event.metaKey) &&
    !event.altKey &&
    !event.shiftKey &&
    !(target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement || target.isContentEditable)
  ) {
    const action = formattingHotkeys[event.key.toLowerCase()];

    if (action) {
      event.preventDefault();
      action();
    }
  }
});

function mergeSelection() {
  const worksheet = activeElement?.worksheet;

  if (worksheet) {
    const { left, top, right, bottom } = getSelection().range;
    // Unmerges only if all selected cells are merged, otherwise merges all of them.
    const unmerge = isSelectionMerged();
    let skipped = false;

    // The selection always covers whole merged cells.
    if (!isEditable(left, top) || !isEditable(right, bottom)) {
      skipped = true;
    } else if (unmerge) {
      worksheet.edit(edit =>
        worksheet
          .getMerges(left, top, right - left + 1, bottom - top + 1)
          .forEach(({ x, y }) => edit.unmergeCells(x, y))
      );

      // Keeps all cells of the unmerged cells selected, the selection may be only the merged cell itself.
      /** @type {WorksheetElement} */ (activeElement).selection = new WorksheetSelection(
        { column: left, row: top },
        { left, top, right, bottom }
      );
    } else {
      worksheet.edit(edit => edit.mergeCells(left, top, right - left + 1, bottom - top + 1));
    }

    skipped && showStatus('Some cells are locked and cannot be merged.');
    updateFormattingButtons();
    updateUndoButtons();
  }

  activeElement?.focus();
}

function updateUndoButtons() {
  const { canUndo, canRedo } = activeElement ?? { canUndo: false, canRedo: false };

  ui.update({ canRedo, canUndo });
}

/** @param {'undo' | 'redo'} action */
function undoOrRedo(action) {
  if (activeElement && activeElement[action]()) {
    updateFormattingButtons();
    updateSelectionInfo();
  }

  activeElement?.focus();
}

document.addEventListener('keydown', event => {
  const target = /** @type {HTMLElement} */ (event.target);

  // Let text fields (e.g. the cell editor) handle their own undo.
  if (
    !(event.ctrlKey || event.metaKey) ||
    event.altKey ||
    target.isContentEditable ||
    target instanceof HTMLInputElement ||
    target instanceof HTMLTextAreaElement
  ) {
    return;
  }

  const key = event.key.toLowerCase();

  if (key === 'z' && !event.shiftKey) {
    event.preventDefault();
    undoOrRedo('undo');
  } else if (key === 'y' || (key === 'z' && event.shiftKey)) {
    event.preventDefault();
    undoOrRedo('redo');
  }
});

addSheetButton.addEventListener('click', () => {
  if (workbook.worksheets.length < MAX_WORKSHEETS) {
    activateWorksheet(workbook.addWorksheet());
  }
});

/** @param {import('./ui/Toolbar.jsx').BorderPreset} preset */
function setBorder(preset) {
  const worksheet = activeElement?.worksheet;

  if (worksheet) {
    const { left, top, right, bottom } = getSelection().range;
    const range = { x: left, y: top, width: right - left + 1, height: bottom - top + 1 };

    worksheet.edit(edit => {
      if (preset === 'all' || preset === 'none') {
        edit.setBorder(range, new Set(['outside', 'inside']), preset === 'all');
      } else {
        edit.setBorder(range, new Set([/** @type {import('./worksheet/Worksheet.js').BorderEdges} */ (preset)]));
      }
    });

    updateUndoButtons();
  }

  activeElement?.focus();
}

/** @param {boolean} nextLocked */
function setLocked(nextLocked) {
  locked = nextLocked;

  for (const worksheet of workbook.worksheets) {
    getElement(worksheet).toggleAttribute('locked', locked);
  }

  ui.update({ locked });
  activeElement?.focus();
}

fileInput.addEventListener('change', async () => {
  const file = fileInput.files?.[0];

  fileInput.value = '';

  if (!file) {
    return;
  }

  try {
    if (/\.zip$/i.test(file.name) || file.type === 'application/zip') {
      if (!confirm('Opening a ZIP file replaces all sheets in the current workbook. Continue?')) {
        return;
      }

      setWorkbook(await Workbook.fromZip(await file.arrayBuffer()));
    } else {
      if (workbook.worksheets.length >= MAX_WORKSHEETS) {
        alert(`A workbook cannot have more than ${MAX_WORKSHEETS} sheets.`);

        return;
      }

      const name = workbook.getUniqueWorksheetName(file.name.replace(/\.csv$/i, ''));

      const text = await file.text();
      const worksheet = workbook.addWorksheet(name);

      worksheet.edit(edit => edit.loadCSV(text));
      activateWorksheet(worksheet);
    }
  } catch (error) {
    alert(`Failed to open "${file.name}": ${/** @type {Error} */ (error).message}`);
  }
});

function saveCSV() {
  const worksheet = activeElement?.worksheet;

  if (worksheet) {
    // Byte order mark so Excel recognizes the file as UTF-8.
    download(
      new Blob(['\ufeff', worksheet.toCSV()], { type: 'text/csv' }),
      `${workbook.getWorksheetName(worksheet)}.csv`
    );
  }
}

async function saveZip() {
  try {
    download(new Blob([await workbook.toZip()], { type: 'application/zip' }), 'workbook.zip');
  } catch (error) {
    alert(`Failed to save workbook: ${/** @type {Error} */ (error).message}`);
  }
}

async function printSelection() {
  const worksheet = activeElement?.worksheet;

  if (!worksheet) {
    return;
  }

  const { left, top, right, bottom } = getSelection().range;
  const rows = worksheet.getValues(left, top, right - left + 1, bottom - top + 1);

  try {
    const getBorder = (/** @type {number} */ column, /** @type {number} */ row) =>
      left + column >= 0 && top + row >= 0 && left + column < worksheet.columnCount && top + row < worksheet.rowCount
        ? worksheet.getBorder(left + column, top + row)
        : undefined;

    const columnWidths = [];

    for (let column = left; column <= right; column++) {
      columnWidths.push(worksheet.getColumnWidth(column));
    }

    const formattings = worksheet.getFormattings(left, top, right - left + 1, bottom - top + 1);
    /** Column and row spans of merged cells, keyed by `row,column`, 0 for cells not in the first column or row. @type {Map<string, { columns: number; rows: number }>} */
    const spans = new Map();

    // The selection always covers whole merged cells.
    for (const merge of worksheet.getMerges(left, top, right - left + 1, bottom - top + 1)) {
      for (let row = merge.y; row < merge.y + merge.height; row++) {
        for (let column = merge.x; column < merge.x + merge.width; column++) {
          spans.set(`${row - top},${column - left}`, {
            columns: column === merge.x ? merge.width : 0,
            rows: row === merge.y ? merge.height : 0
          });
        }
      }
    }

    const getColumnSpan = (/** @type {number} */ column, /** @type {number} */ row) =>
      spans.get(`${row},${column}`)?.columns ?? 1;
    const getRowSpan = (/** @type {number} */ column, /** @type {number} */ row) =>
      spans.get(`${row},${column}`)?.rows ?? 1;
    const getFormatting = (/** @type {number} */ column, /** @type {number} */ row) => formattings[row]?.[column];

    download(
      await cellsToPNG(rows, { columnWidths, getBorder, getColumnSpan, getFormatting, getRowSpan }),
      `${workbook.getWorksheetName(worksheet)}.png`
    );
  } catch (error) {
    alert(`Failed to print selected cells: ${/** @type {Error} */ (error).message}`);
  }
}

setWorkbook(new Workbook());
