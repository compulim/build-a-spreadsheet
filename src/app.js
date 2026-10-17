import { SpreadsheetGrid } from './grid.js';
import { formatValue } from './csv.js';
import { HIGHLIGHTS } from './highlights.js';
import { cellsToPNG } from './printCells.js';
import { MAX_WORKSHEETS } from './workbook/constants.js';
import { getColumnName } from './workbook/getColumnName.js';
import { Workbook } from './workbook/Workbook.js';
import { Worksheet } from './workbook/Worksheet.js';

/**
 * Number of leading rows and columns that become read-only, and frozen on screen, when the lock button is pressed.
 * Change these to protect a different range, e.g. 2 header rows and 1 ID column.
 */
const LOCKED_ROW_COUNT = 1;
const LOCKED_COLUMN_COUNT = 1;

const gridContainer = /** @type {HTMLElement} */ (document.getElementById('grid-container'));
const sheetTabs = /** @type {HTMLElement} */ (document.getElementById('sheet-tabs'));
const addSheetButton = /** @type {HTMLButtonElement} */ (document.getElementById('add-sheet-button'));
const openButton = /** @type {HTMLButtonElement} */ (document.getElementById('open-button'));
const fileInput = /** @type {HTMLInputElement} */ (document.getElementById('file-input'));
const saveCSVButton = /** @type {HTMLButtonElement} */ (document.getElementById('save-csv-button'));
const saveZipButton = /** @type {HTMLButtonElement} */ (document.getElementById('save-zip-button'));
const printButton = /** @type {HTMLButtonElement} */ (document.getElementById('print-button'));
const undoButton = /** @type {HTMLButtonElement} */ (document.getElementById('undo-button'));
const redoButton = /** @type {HTMLButtonElement} */ (document.getElementById('redo-button'));
const borderSelect = /** @type {HTMLSelectElement} */ (document.getElementById('border-select'));
const lockButton = /** @type {HTMLButtonElement} */ (document.getElementById('lock-button'));
const unformatButton = /** @type {HTMLButtonElement} */ (document.getElementById('unformat-button'));
const mergeButton = /** @type {HTMLButtonElement} */ (document.getElementById('merge-button'));
const styleButtons = /** @type {const} */ (['bold', 'italic', 'underline', 'strikethrough']).map(style => ({
  style,
  button: /** @type {HTMLButtonElement} */ (document.getElementById(`${style}-button`))
}));
const alignButtons = /** @type {const} */ (['left', 'center', 'right']).map(align => ({
  align,
  button: /** @type {HTMLButtonElement} */ (document.getElementById(`align-${align}-button`))
}));
const formatButtons = /** @type {const} */ (['currency', 'percent', 'date', 'datetime', 'duration']).map(format => ({
  format,
  button: /** @type {HTMLButtonElement} */ (document.getElementById(`format-${format}-button`))
}));
const highlightButtons = HIGHLIGHTS.map(({ name, label, background, color }) => {
  const button = document.createElement('button');

  button.type = 'button';
  button.className = 'toolbar__highlight-button';
  button.title = label;
  const icon = document.createElement('span');

  icon.className = 'grid__cell-text--marker';
  icon.textContent = 'A';
  button.append(icon);
  button.setAttribute('aria-label', label);
  button.setAttribute('aria-pressed', 'false');
  background && button.style.setProperty('--marker-color', background);
  color && button.style.setProperty('--highlight-color', color);
  unformatButton.before(button);

  return { name, button };
});
const nameBox = /** @type {HTMLOutputElement} */ (document.getElementById('name-box'));
const valueBox = /** @type {HTMLOutputElement} */ (document.getElementById('value-box'));
const status = /** @type {HTMLOutputElement} */ (document.getElementById('status'));

const grid = new SpreadsheetGrid(gridContainer, {
  lockedColumnCount: LOCKED_COLUMN_COUNT,
  lockedRowCount: LOCKED_ROW_COUNT
});

/** @type {Workbook} */
let workbook;

/** @type {ReturnType<typeof setTimeout> | undefined} */
let statusTimeout;

/** @param {string} message */
function showStatus(message) {
  clearTimeout(statusTimeout);
  status.value = message;
  statusTimeout = setTimeout(() => (status.value = ''), 4000);
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
  const active = grid.worksheet;
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
  grid.setWorksheet(worksheet);
  renderTabs();
  updateUndoButtons();
  updateFormattingButtons();
  grid.focus();
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

  if (grid.worksheet === worksheet) {
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
    workbook.removeEventListener('worksheetrename', renderTabs);
    workbook.removeEventListener('cellchange', updateUndoButtons);
    workbook.removeEventListener('cellchange', updateFormattingButtons);
  }

  workbook = nextWorkbook;
  workbook.addEventListener('worksheetadd', renderTabs);
  workbook.addEventListener('worksheetremove', renderTabs);
  workbook.addEventListener('worksheetrename', renderTabs);
  workbook.addEventListener('cellchange', updateUndoButtons);
  workbook.addEventListener('cellchange', updateFormattingButtons);

  // Expose for scripting from the browser console, e.g. `app.workbook.worksheets[0].setValue(0, 0, 'Hello')`.
  Object.assign(globalThis, { app: { workbook, grid } });

  grid.setWorksheet(undefined);
  activateWorksheet(workbook.worksheets[0]);
}

function updateSelectionInfo() {
  const { left, top, right, bottom } = grid.selection;
  const { column, row } = grid.activeCell;
  const topLeft = `${getColumnName(left)}${top + 1}`;

  nameBox.value = left === right && top === bottom ? topLeft : `${topLeft}:${getColumnName(right)}${bottom + 1}`;
  valueBox.value = grid.worksheet ? formatValue(grid.worksheet.getValue(column, row)) : '';
  updateFormattingButtons();
}

function updateFormattingButtons() {
  const { column, row } = grid.activeCell;
  const formatting = grid.worksheet?.getFormattings(column, row, 1, 1)[0][0];

  for (const { style, button } of styleButtons) {
    button.setAttribute('aria-pressed', String(!!formatting?.[style]));
  }

  for (const { align, button } of alignButtons) {
    button.setAttribute('aria-pressed', String(formatting?.align === align));
  }

  for (const { format, button } of formatButtons) {
    button.setAttribute('aria-pressed', String(formatting?.format === format));
  }

  for (const { name, button } of highlightButtons) {
    button.setAttribute('aria-pressed', String(formatting?.highlight === name));
  }

  mergeButton.setAttribute('aria-pressed', String(isSelectionMerged()));
}

/** Whether every selected cell is in a merged cell. */
function isSelectionMerged() {
  const worksheet = grid.worksheet;

  if (!worksheet) {
    return false;
  }

  const { left, top, right, bottom } = grid.selection;

  for (let row = top; row <= bottom; row++) {
    for (let column = left; column <= right; column++) {
      if (!worksheet.getMerge(column, row)) {
        return false;
      }
    }
  }

  return true;
}

/** @param {import('./workbook/Worksheet.js').FormattingUpdater} updater */
function formatSelection(updater) {
  const worksheet = grid.worksheet;

  if (worksheet) {
    const { left, top, right, bottom } = grid.selection;

    worksheet.setFormattings(updater, left, top, right - left + 1, bottom - top + 1);
    updateFormattingButtons();
    updateUndoButtons();
  }

  grid.focus();
}

for (const { style, button } of styleButtons) {
  button.addEventListener('click', () => {
    const value = button.getAttribute('aria-pressed') !== 'true';

    formatSelection(formatting => ({ ...formatting, [style]: value }));
  });
}

for (const { align, button } of alignButtons) {
  button.addEventListener('click', () => {
    // Clicking a pushed align button unpushes it, back to the default (type-based) alignment.
    const nextAlign = button.getAttribute('aria-pressed') === 'true' ? undefined : align;

    formatSelection(formatting => ({ ...formatting, align: nextAlign }));
  });
}

for (const { format, button } of formatButtons) {
  button.addEventListener('click', () => {
    // Only one number format at a time, clicking a pushed format button unpushes it, back to the default format.
    const nextFormat = button.getAttribute('aria-pressed') === 'true' ? undefined : format;

    formatSelection(formatting => ({ ...formatting, format: nextFormat }));
  });
}

for (const { name, button } of highlightButtons) {
  button.addEventListener('click', () => {
    // Only one highlight at a time, clicking a pushed highlight button unpushes it.
    const nextHighlight = button.getAttribute('aria-pressed') === 'true' ? undefined : name;

    formatSelection(formatting => ({ ...formatting, highlight: nextHighlight }));
  });
}

/** @type {Record<string, HTMLButtonElement>} */
const formattingHotkeys = {
  b: styleButtons[0].button,
  i: styleButtons[1].button,
  u: styleButtons[2].button,
  l: alignButtons[0].button,
  e: alignButtons[1].button,
  r: alignButtons[2].button
};

gridContainer.addEventListener('keydown', event => {
  const target = /** @type {HTMLElement} */ (event.target);

  if (
    (event.ctrlKey || event.metaKey) &&
    !event.altKey &&
    !event.shiftKey &&
    !(target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement || target.isContentEditable)
  ) {
    const button = formattingHotkeys[event.key.toLowerCase()];

    if (button) {
      event.preventDefault();
      button.click();
    }
  }
});

unformatButton.addEventListener('click', () => formatSelection(() => null));

mergeButton.addEventListener('click', () => {
  const worksheet = grid.worksheet;

  if (worksheet) {
    const { left, top, right, bottom } = grid.selection;
    // Unmerges only if all selected cells are merged, otherwise merges all of them.
    const unmerge = isSelectionMerged();
    let skipped = false;

    // Cells are merged horizontally only, each selected row separately.
    for (let row = top; row <= bottom; row++) {
      const merges = worksheet.getMerges(row).filter(({ x, width }) => x <= right && x + width - 1 >= left);
      const start = Math.min(left, ...merges.map(({ x }) => x));
      const end = Math.max(right, ...merges.map(({ x, width }) => x + width - 1));

      if (!grid.isEditable(start, row) || !grid.isEditable(end, row)) {
        skipped = true;
      } else if (unmerge) {
        merges.forEach(({ x }) => worksheet.unmergeCells(x, row));
      } else if (end > start) {
        worksheet.mergeCells(start, row, end - start + 1);
      }
    }

    skipped && showStatus('Some cells are locked and cannot be merged.');
    updateFormattingButtons();
    updateUndoButtons();
  }

  grid.focus();
});

function updateUndoButtons() {
  undoButton.disabled = !grid.worksheet?.canUndo;
  redoButton.disabled = !grid.worksheet?.canRedo;
}

/** @param {'undo' | 'redo'} action */
function undoOrRedo(action) {
  const worksheet = grid.worksheet;

  if (!worksheet) {
    return;
  }

  /** @type {{ x: number; y: number }[]} */
  const changes = [];
  const handleCellChange = (/** @type {Event} */ event) =>
    changes.push(/** @type {import('./workbook/private/CellChangeEvent.js').CellChangeEvent} */ (event));

  worksheet.addEventListener('cellchange', handleCellChange);
  worksheet.addEventListener('formattingchange', handleCellChange);

  try {
    if (!worksheet[action]()) {
      return;
    }

    if (grid.locked && changes.some(({ x, y }) => x < LOCKED_COLUMN_COUNT || y < LOCKED_ROW_COUNT)) {
      // Revert, locked cells cannot be changed.
      worksheet[action === 'undo' ? 'redo' : 'undo']();
      showStatus(`Cannot ${action} because some cells are locked.`);

      return;
    }
  } finally {
    worksheet.removeEventListener('cellchange', handleCellChange);
    worksheet.removeEventListener('formattingchange', handleCellChange);
    updateUndoButtons();
    updateFormattingButtons();
  }

  grid.focus();

  if (changes.length) {
    const xs = changes.map(({ x }) => x);
    const ys = changes.map(({ y }) => y);

    grid.select({ column: Math.min(...xs), row: Math.min(...ys) }, { column: Math.max(...xs), row: Math.max(...ys) });
  }

  updateSelectionInfo();
}

undoButton.addEventListener('click', () => undoOrRedo('undo'));
redoButton.addEventListener('click', () => undoOrRedo('redo'));

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

grid.addEventListener('selectionchange', updateSelectionInfo);
grid.addEventListener('lockedit', () => showStatus('Some cells are locked and cannot be edited.'));

addSheetButton.addEventListener('click', () => {
  if (workbook.worksheets.length < MAX_WORKSHEETS) {
    activateWorksheet(workbook.addWorksheet());
  }
});

borderSelect.addEventListener('change', () => {
  const preset = borderSelect.value;
  const worksheet = grid.worksheet;

  borderSelect.value = '';

  if (worksheet) {
    const { left, top, right, bottom } = grid.selection;
    const range = { x: left, y: top, width: right - left + 1, height: bottom - top + 1 };

    if (preset === 'all' || preset === 'none') {
      worksheet.setBorder(range, new Set(['outside', 'inside']), preset === 'all');
    } else {
      worksheet.setBorder(range, new Set([/** @type {import('./workbook/Worksheet.js').BorderEdges} */ (preset)]));
    }
  }

  grid.focus();
});

lockButton.addEventListener('click', () => {
  grid.locked = !grid.locked;
  lockButton.setAttribute('aria-pressed', String(grid.locked));
  lockButton.textContent = grid.locked ? '🔒 Locked' : '🔓 Unlocked';
  grid.focus();
});

openButton.addEventListener('click', () => fileInput.click());

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

      activateWorksheet(workbook.addWorksheet(name, Worksheet.fromCSV(await file.text())));
    }
  } catch (error) {
    alert(`Failed to open "${file.name}": ${/** @type {Error} */ (error).message}`);
  }
});

saveCSVButton.addEventListener('click', () => {
  const worksheet = grid.worksheet;

  if (worksheet) {
    // Byte order mark so Excel recognizes the file as UTF-8.
    download(new Blob(['\ufeff', worksheet.toCSV()], { type: 'text/csv' }), `${workbook.getWorksheetName(worksheet)}.csv`);
  }
});

saveZipButton.addEventListener('click', async () => {
  try {
    download(new Blob([await workbook.toZip()], { type: 'application/zip' }), 'workbook.zip');
  } catch (error) {
    alert(`Failed to save workbook: ${/** @type {Error} */ (error).message}`);
  }
});

printButton.addEventListener('click', async () => {
  const worksheet = grid.worksheet;

  if (!worksheet) {
    return;
  }

  const { left, top, right, bottom } = grid.selection;
  const rows = worksheet.getValues(left, top, right - left + 1, bottom - top + 1);

  try {
    const getBorder = (/** @type {number} */ column, /** @type {number} */ row) =>
      left + column >= 0 &&
      top + row >= 0 &&
      left + column < worksheet.columnCount &&
      top + row < worksheet.rowCount
        ? worksheet.getBorder(left + column, top + row)
        : undefined;

    const columnWidths = [];

    for (let column = left; column <= right; column++) {
      columnWidths.push(worksheet.getColumnWidth(column));
    }

    const formattings = worksheet.getFormattings(left, top, right - left + 1, bottom - top + 1);
    /** Column spans of merged cells (clipped to the selection), keyed by `row,column`, 0 for covered cells. @type {Map<string, number>} */
    const spans = new Map();

    for (let row = top; row <= bottom; row++) {
      for (const merge of worksheet.getMerges(row)) {
        const start = Math.max(merge.x, left);
        const end = Math.min(merge.x + merge.width - 1, right);

        if (start > end) {
          continue;
        }

        // The selection may start in the middle of a merged cell, whose content and formatting are in its first cell.
        rows[row - top][start - left] = worksheet.getValue(merge.x, row) ?? '';
        formattings[row - top][start - left] = worksheet.getFormattings(merge.x, row, 1, 1)[0][0];

        for (let column = start; column <= end; column++) {
          spans.set(`${row - top},${column - left}`, column === start ? end - start + 1 : 0);
        }
      }
    }

    const getColumnSpan = (/** @type {number} */ column, /** @type {number} */ row) => spans.get(`${row},${column}`) ?? 1;
    const getFormatting = (/** @type {number} */ column, /** @type {number} */ row) => formattings[row]?.[column];

    download(await cellsToPNG(rows, { columnWidths, getBorder, getColumnSpan, getFormatting }), `${workbook.getWorksheetName(worksheet)}.png`);
  } catch (error) {
    alert(`Failed to print selected cells: ${/** @type {Error} */ (error).message}`);
  }
});

setWorkbook(new Workbook());
