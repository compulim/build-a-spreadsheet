import { formatValue } from './csv.js';
import { HIGHLIGHTS } from './highlights.js';
import { cellsToPNG } from './printCells.js';
import { MAX_WORKSHEETS } from './workbook/constants.js';
import { getColumnName } from './workbook/getColumnName.js';
import { Workbook } from './workbook/Workbook.js';
import {
  getWorksheetHistory,
  getWorksheetSelection,
  redoWorksheet,
  selectWorksheetCells,
  undoWorksheet,
  WorksheetElement
} from './WorksheetElement.js';

/** @typedef {import('./workbook/Worksheet.js').Worksheet} Worksheet */

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
  mountWorksheet(/** @type {import('./workbook/private/WorksheetEvent.js').WorksheetEvent} */ (event).worksheet);
}

/** Active cell and selection of the active worksheet. */
function getSelection() {
  return activeElement
    ? getWorksheetSelection(activeElement)
    : { activeCell: { column: 0, row: 0 }, selection: { left: 0, top: 0, right: 0, bottom: 0 } };
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
    workbook.removeEventListener('cellchange', updateUndoButtons);
    workbook.removeEventListener('cellchange', updateFormattingButtons);
  }

  workbook = nextWorkbook;
  workbook.addEventListener('worksheetadd', renderTabs);
  workbook.addEventListener('worksheetremove', renderTabs);
  workbook.addEventListener('worksheetadd', handleWorksheetAdd);
  workbook.addEventListener('worksheetrename', renderTabs);
  workbook.addEventListener('cellchange', updateUndoButtons);
  workbook.addEventListener('cellchange', updateFormattingButtons);

  // Expose for scripting from the browser console, e.g. `app.workbook.worksheets[0].setValue(0, 0, 'Hello')`.
  Object.assign(globalThis, { app: { workbook } });

  gridContainer.replaceChildren();
  workbook.worksheets.forEach(mountWorksheet);
  activateWorksheet(workbook.worksheets[0]);
}

function updateSelectionInfo() {
  const { left, top, right, bottom } = getSelection().selection;
  const { column, row } = getSelection().activeCell;
  const topLeft = `${getColumnName(left)}${top + 1}`;

  nameBox.value = left === right && top === bottom ? topLeft : `${topLeft}:${getColumnName(right)}${bottom + 1}`;
  valueBox.value = activeElement ? formatValue(activeElement.worksheet.getValue(column, row)) : '';
  updateFormattingButtons();
}

function updateFormattingButtons() {
  const { column, row } = getSelection().activeCell;
  const formatting = activeElement?.worksheet?.getFormattings(column, row, 1, 1)[0][0];

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
  const worksheet = activeElement?.worksheet;

  if (!worksheet) {
    return false;
  }

  const { left, top, right, bottom } = getSelection().selection;

  let covered = 0;

  for (const { x, y, width, height } of worksheet.getMerges(left, top, right - left + 1, bottom - top + 1)) {
    covered +=
      (Math.min(x + width - 1, right) - Math.max(x, left) + 1) * (Math.min(y + height - 1, bottom) - Math.max(y, top) + 1);
  }

  // Merged cells never overlap, so all cells are merged if they cover the whole selection.
  return covered === (right - left + 1) * (bottom - top + 1);
}

/** @param {import('./workbook/Worksheet.js').FormattingUpdater} updater */
function formatSelection(updater) {
  const worksheet = activeElement?.worksheet;

  if (worksheet) {
    const { left, top, right, bottom } = getSelection().selection;

    worksheet.setFormattings(updater, left, top, right - left + 1, bottom - top + 1);
    updateFormattingButtons();
    updateUndoButtons();
  }

  activeElement?.focus();
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
  r: alignButtons[2].button,
  m: mergeButton
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
  const worksheet = activeElement?.worksheet;

  if (worksheet) {
    const { left, top, right, bottom } = getSelection().selection;
    // Unmerges only if all selected cells are merged, otherwise merges all of them.
    const unmerge = isSelectionMerged();
    let skipped = false;

    // The selection always covers whole merged cells.
    if (!isEditable(left, top) || !isEditable(right, bottom)) {
      skipped = true;
    } else if (unmerge) {
      worksheet
        .getMerges(left, top, right - left + 1, bottom - top + 1)
        .forEach(({ x, y }) => worksheet.unmergeCells(x, y));

      // Keeps all cells of the unmerged cells selected, the selection may be only the merged cell itself.
      selectWorksheetCells(/** @type {WorksheetElement} */ (activeElement), { column: left, row: top }, { column: right, row: bottom });
    } else {
      worksheet.mergeCells(left, top, right - left + 1, bottom - top + 1);
    }

    skipped && showStatus('Some cells are locked and cannot be merged.');
    updateFormattingButtons();
    updateUndoButtons();
  }

  activeElement?.focus();
});

function updateUndoButtons() {
  const { canUndo, canRedo } = activeElement ? getWorksheetHistory(activeElement) : { canUndo: false, canRedo: false };

  undoButton.disabled = !canUndo;
  redoButton.disabled = !canRedo;
}

/** @param {'undo' | 'redo'} action */
function undoOrRedo(action) {
  if (activeElement && (action === 'undo' ? undoWorksheet : redoWorksheet)(activeElement)) {
    updateFormattingButtons();
    updateSelectionInfo();
  }

  activeElement?.focus();
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


addSheetButton.addEventListener('click', () => {
  if (workbook.worksheets.length < MAX_WORKSHEETS) {
    activateWorksheet(workbook.addWorksheet());
  }
});

borderSelect.addEventListener('change', () => {
  const preset = borderSelect.value;
  const worksheet = activeElement?.worksheet;

  borderSelect.value = '';

  if (worksheet) {
    const { left, top, right, bottom } = getSelection().selection;
    const range = { x: left, y: top, width: right - left + 1, height: bottom - top + 1 };

    if (preset === 'all' || preset === 'none') {
      worksheet.setBorder(range, new Set(['outside', 'inside']), preset === 'all');
    } else {
      worksheet.setBorder(range, new Set([/** @type {import('./workbook/Worksheet.js').BorderEdges} */ (preset)]));
    }

    updateUndoButtons();
  }

  activeElement?.focus();
});

lockButton.addEventListener('click', () => {
  locked = !locked;

  for (const worksheet of workbook.worksheets) {
    getElement(worksheet).toggleAttribute('locked', locked);
  }

  lockButton.setAttribute('aria-pressed', String(locked));
  lockButton.textContent = locked ? '🔒 Locked' : '🔓 Unlocked';
  activeElement?.focus();
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

      const text = await file.text();
      const worksheet = workbook.addWorksheet(name);

      worksheet.loadCSV(text);
      activateWorksheet(worksheet);
    }
  } catch (error) {
    alert(`Failed to open "${file.name}": ${/** @type {Error} */ (error).message}`);
  }
});

saveCSVButton.addEventListener('click', () => {
  const worksheet = activeElement?.worksheet;

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
  const worksheet = activeElement?.worksheet;

  if (!worksheet) {
    return;
  }

  const { left, top, right, bottom } = getSelection().selection;
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
    const getRowSpan = (/** @type {number} */ column, /** @type {number} */ row) => spans.get(`${row},${column}`)?.rows ?? 1;
    const getFormatting = (/** @type {number} */ column, /** @type {number} */ row) => formattings[row]?.[column];

    download(await cellsToPNG(rows, { columnWidths, getBorder, getColumnSpan, getFormatting, getRowSpan }), `${workbook.getWorksheetName(worksheet)}.png`);
  } catch (error) {
    alert(`Failed to print selected cells: ${/** @type {Error} */ (error).message}`);
  }
});

setWorkbook(new Workbook());
