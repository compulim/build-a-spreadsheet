import { SpreadsheetGrid } from './grid.js';
import { formatValue } from './csv.js';
import { cellsToPNG } from './printCells.js';
import { UndoManager } from './UndoManager.js';
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
const lockButton = /** @type {HTMLButtonElement} */ (document.getElementById('lock-button'));
const nameBox = /** @type {HTMLOutputElement} */ (document.getElementById('name-box'));
const valueBox = /** @type {HTMLOutputElement} */ (document.getElementById('value-box'));
const status = /** @type {HTMLOutputElement} */ (document.getElementById('status'));

const grid = new SpreadsheetGrid(gridContainer, {
  lockedColumnCount: LOCKED_COLUMN_COUNT,
  lockedRowCount: LOCKED_ROW_COUNT
});

const undoManager = new UndoManager();

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
      label.textContent = worksheet.name;
      tab.append(label);

      tab.addEventListener('click', () => activateWorksheet(worksheet));
      tab.addEventListener('dblclick', () => renameWorksheet(worksheet));

      if (worksheets.length > 1) {
        const close = document.createElement('button');

        close.type = 'button';
        close.className = 'sheet-tab__close';
        close.textContent = '×';
        close.title = `Delete "${worksheet.name}"`;
        close.setAttribute('aria-label', `Delete sheet ${worksheet.name}`);
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
  grid.focus();
}

/** @param {Worksheet} worksheet */
function renameWorksheet(worksheet) {
  const name = prompt('Rename sheet', worksheet.name);

  if (name === null) {
    return;
  }

  try {
    worksheet.name = name.trim();
  } catch (error) {
    alert(/** @type {Error} */ (error).message);
  }
}

/** @param {Worksheet} worksheet */
function deleteWorksheet(worksheet) {
  if (workbook.worksheets.length <= 1 || !confirm(`Delete sheet "${worksheet.name}"?`)) {
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
    workbook.removeEventListener('worksheetremove', resetUndo);
    undoManager.untrack(workbook);
  }

  workbook = nextWorkbook;
  workbook.addEventListener('worksheetadd', renderTabs);
  workbook.addEventListener('worksheetremove', renderTabs);
  workbook.addEventListener('worksheetrename', renderTabs);
  // Undo history may refer to cells of a removed sheet, forget it.
  workbook.addEventListener('worksheetremove', resetUndo);
  undoManager.track(workbook);
  undoManager.reset();

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
}

function resetUndo() {
  undoManager.reset();
}

/** @param {'undo' | 'redo'} action */
function undoOrRedo(action) {
  const next = action === 'undo' ? undoManager.nextUndo : undoManager.nextRedo;

  if (grid.locked && next?.some(({ column, row }) => column < LOCKED_COLUMN_COUNT || row < LOCKED_ROW_COUNT)) {
    showStatus(`Cannot ${action} because some cells are locked.`);

    return;
  }

  const changes = undoManager[action]();

  if (!changes?.length) {
    return;
  }

  const { worksheet } = changes[0];

  grid.worksheet === worksheet ? grid.focus() : activateWorksheet(worksheet);

  const columns = changes.map(({ column }) => column);
  const rows = changes.map(({ row }) => row);

  grid.select(
    { column: Math.min(...columns), row: Math.min(...rows) },
    { column: Math.max(...columns), row: Math.max(...rows) }
  );
  updateSelectionInfo();
}

undoManager.addEventListener('change', () => {
  undoButton.disabled = !undoManager.canUndo;
  redoButton.disabled = !undoManager.canRedo;
});

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

      activateWorksheet(workbook.addWorksheet(Worksheet.fromCSV(await file.text(), name)));
    }
  } catch (error) {
    alert(`Failed to open "${file.name}": ${/** @type {Error} */ (error).message}`);
  }
});

saveCSVButton.addEventListener('click', () => {
  const worksheet = grid.worksheet;

  if (worksheet) {
    // Byte order mark so Excel recognizes the file as UTF-8.
    download(new Blob(['\ufeff', worksheet.toCSV()], { type: 'text/csv' }), `${worksheet.name}.csv`);
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
  const rows = [];

  for (let row = top; row <= bottom; row++) {
    const values = [];

    for (let column = left; column <= right; column++) {
      values.push(worksheet.getValue(column, row));
    }

    rows.push(values);
  }

  try {
    download(await cellsToPNG(rows), `${worksheet.name}.png`);
  } catch (error) {
    alert(`Failed to print selected cells: ${/** @type {Error} */ (error).message}`);
  }
});

setWorkbook(new Workbook());
