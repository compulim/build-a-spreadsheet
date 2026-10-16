# build-a-spreadsheet

A high-performance spreadsheet web app that runs entirely in the browser — no server, no build step, no dependencies.
Written in plain HTML, CSS and JavaScript (ES modules), so it can be served as static files from GitHub Pages or any CDN.

## Features

- 50 columns × 10,000 rows per sheet, with column (A, B, … AX) and row headers.
- Resize a column by dragging the right border of its header; double-click the border to auto-fit the column to its content. When multiple whole columns are selected, all of them are resized to the same width.
- Virtualized rendering: only the visible rows are in the DOM, while scrolling uses the browser's native scrollbars.
- Smooth scrolling: rows are rendered synchronously on scroll with a viewport of pre-rendered rows on each side, so wheel, touchpad and momentum scrolling stay seamless. Keyboard navigation scrolls smoothly for short distances (instant for long jumps, or when the user prefers reduced motion).
- Keyboard navigation: arrow keys, <kbd>Ctrl</kbd>+arrow (jump to edge), <kbd>Home</kbd>/<kbd>End</kbd>, <kbd>Page Up</kbd>/<kbd>Page Down</kbd>, <kbd>Enter</kbd>/<kbd>Tab</kbd>.
- Multiple selection with <kbd>Shift</kbd>+click, <kbd>Shift</kbd>+arrow keys, mouse drag, or by clicking row/column headers; <kbd>Ctrl</kbd>+<kbd>A</kbd> selects all.
- Editing: type to replace, <kbd>F2</kbd> or double-click to edit, <kbd>Escape</kbd> to cancel, <kbd>Delete</kbd> to clear. A leading `'` keeps the value as text.
- Copy, cut and paste with Excel and Excel Web App (tab-separated text and HTML table on the clipboard). Pasting a single value into a larger selection fills it.
- Values are booleans, numbers or strings. Formulas are not supported.
- Up to 50 named sheets in the bottom bar: <kbd>+</kbd> to add, double-click to rename, × to delete.
- Lock button: makes the rows/columns listed in `LOCKED_ROWS` and `LOCKED_COLUMNS` in [`src/app.js`](src/app.js) read-only.
- Open a CSV file (added as a new sheet) or a ZIP file of CSV files (replaces the workbook). Save the current sheet as CSV, or the whole workbook as a ZIP of CSV files.

## Running locally

ES modules must be served over HTTP:

```sh
npx serve .
# or
python3 -m http.server
```

To deploy, publish the repository root (e.g. enable GitHub Pages for the branch). All paths are relative.

## Programming interface

```js
import { Workbook } from './src/workbook/Workbook.js';
import { Worksheet } from './src/workbook/Worksheet.js';

const workbook = new Workbook(); // EventTarget
const sheet = workbook.addWorksheet('Data'); // Worksheet, also an EventTarget

workbook.addEventListener('cellchange', ({ worksheet, column, row, value, previousValue }) => {
  console.log(worksheet.name, column, row, previousValue, '→', value);
});

sheet.setValue(0, 0, 'Name'); // setValue(column: number, row: number, value: boolean | number | string)
sheet.setValue(1, 0, 42);
sheet.setValue(2, 0, true);
sheet.getValues(); // [['Name', 42, true]] — from A1 to the last used cell, empty cells are ''

sheet.toCSV(); // 'Name,42,TRUE\r\n'
Worksheet.fromCSV('a,1\r\n', 'Imported');

const zip = await workbook.toZip(); // Uint8Array, one "<sheet name>.csv" per worksheet
const loaded = await Workbook.fromZip(zip);
```

| Member | Description |
| --- | --- |
| `Worksheet#setValue(column, row, value)` | Sets a cell (zero-based). `''`, `null` or `undefined` clears it. Throws `RangeError` / `TypeError` on invalid input. |
| `Worksheet#getValue(column, row)` | Returns a cell value, `''` if empty. |
| `Worksheet#getValues()` | Returns a 2D array (`values[row][column]`) from A1 to the last used cell. |
| `Worksheet#setValues(values, column?, row?)` | Sets a block of values. Values outside of the sheet are ignored. |
| `Worksheet#toCSV()` / `Worksheet#loadCSV(text)` / `Worksheet.fromCSV(text, name?)` | CSV serialization. |
| `Workbook#addWorksheet(nameOrWorksheet?, index?)` / `Workbook#removeWorksheet(worksheetOrIndex)` | Manage up to 50 sheets. |
| `Workbook#worksheets` / `Workbook#getWorksheet(nameOrIndex)` | Access sheets. |
| `Workbook#toZip()` / `Workbook.fromZip(data)` | ZIP of CSV files. |
| `cellchange` event | Dispatched by `Worksheet` and re-dispatched by `Workbook` (`CellChangeEvent`). |
| `worksheetadd` / `worksheetremove` / `worksheetrename` events | Dispatched by `Workbook` (`WorksheetEvent`). |

In the browser console, the running app is available as `app.workbook` and `app.grid`.

## Tests

```sh
npm test
```
