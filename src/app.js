import { SpreadsheetGrid } from './grid.js';
import { formatValue } from './csv.js';
import { getColumnName, MAX_WORKSHEETS, Workbook, Worksheet } from './workbook.js';

/**
 * Rows and columns (zero-based) that become read-only when the lock button is pressed.
 * Change these to protect different rows/columns, e.g. a header row and an ID column.
 */
const LOCKED_ROWS = [0];
const LOCKED_COLUMNS = [0];

const gridContainer = /** @type {HTMLElement} */ (document.getElementById('grid-container'));
const sheetTabs = /** @type {HTMLElement} */ (document.getElementById('sheet-tabs'));
const addSheetButton = /** @type {HTMLButtonElement} */ (document.getElementById('add-sheet-button'));
const openButton = /** @type {HTMLButtonElement} */ (document.getElementById('open-button'));
const fileInput = /** @type {HTMLInputElement} */ (document.getElementById('file-input'));
const saveCSVButton = /** @type {HTMLButtonElement} */ (document.getElementById('save-csv-button'));
const saveZipButton = /** @type {HTMLButtonElement} */ (document.getElementById('save-zip-button'));
const lockButton = /** @type {HTMLButtonElement} */ (document.getElementById('lock-button'));
const nameBox = /** @type {HTMLOutputElement} */ (document.getElementById('name-box'));
const valueBox = /** @type {HTMLOutputElement} */ (document.getElementById('value-box'));
const status = /** @type {HTMLOutputElement} */ (document.getElementById('status'));

const grid = new SpreadsheetGrid(gridContainer, { lockedColumns: LOCKED_COLUMNS, lockedRows: LOCKED_ROWS });

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

  workbook = nextWorkbook;
  workbook.addEventListener('worksheetadd', renderTabs);
  workbook.addEventListener('worksheetremove', renderTabs);
  workbook.addEventListener('worksheetrename', renderTabs);

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
  download(new Blob([await workbook.toZip()], { type: 'application/zip' }), 'workbook.zip');
});

setWorkbook(new Workbook());
