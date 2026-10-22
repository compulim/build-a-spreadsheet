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

workbook.addEventListener('cellchange', ({ worksheet, x, y, value, previousValue }) => {
  console.log(workbook.getWorksheetName(worksheet), x, y, previousValue, '→', value);
});

sheet.setValue(0, 0, 'Name'); // setValue(x: number, y: number, value: boolean | number | string), zero-based
sheet.setValue(1, 0, 42);
sheet.setValue(2, 0, true);
sheet.getValues(0, 0, 3, 1); // [['Name', 42, true]] — getValues(x, y, width, height), empty cells are ''
sheet.setValues([['', '']], 1, 0); // Clears B1:C1
sheet.undo(); // true, B1:C1 restored; returns false when there is nothing to undo
sheet.setColumnWidth(0, 150);

sheet.toCSV(); // 'Name,42,TRUE\r\n'
workbook.addWorksheet('Imported').loadCSV('a,1\r\n');

const zip = await workbook.toZip(); // Uint8Array, one "<sheet name>.csv" per worksheet
const loaded = await Workbook.fromZip(zip);
```

In a browser where `<compulim-worksheet>` ([`src/WorksheetElement.js`](src/WorksheetElement.js)) is defined, the workbook instantiates one element per worksheet. The element renders the grid, handles user interactions, and wholly owns the worksheet. Its only public member is the read-only `worksheet` property:

```js
import './src/WorksheetElement.js';

const worksheet = workbook.addWorksheet('Hello');
const element = workbook.getWorksheetElement(worksheet); // <compulim-worksheet>, element.worksheet === worksheet

element.setAttribute('locked-row-count', '1'); // Read when first connected
document.body.append(element);
element.toggleAttribute('locked', true); // Locks row 1
worksheet.setValue(0, 1, 'Hello'); // The grid updates
```

| Member | Description |
| --- | --- |
| `Worksheet#setValue(x, y, valueOrUpdater)` | Sets a cell (zero-based), or updates it with `(value, x, y) => nextValue`. Strings are trimmed. Blank strings, `null` or `undefined` clear it. Throws `RangeError` / `TypeError` on invalid input. |
| `Worksheet#getValue(x, y)` | Returns a cell value, `undefined` if empty. |
| `Worksheet#getValues(x, y, width, height)` | Returns a 2D array (`values[y][x]`) of a range. |
| `Worksheet#setValues(values, x?, y?)` / `Worksheet#setValues(updater, x, y, width, height)` | Sets a block of values (use `''` to clear), or updates every cell of a range with `(value, x, y) => nextValue`. Values outside of the sheet are ignored. |
| `Worksheet#getColumnWidth(x)` / `Worksheet#setColumnWidth(x, width)` | Column widths in pixels, dispatches `columnwidthchange`. |
| `Worksheet#undo()` / `Worksheet#redo()` | Undo/redo cell changes, returns `false` at the end of the stack. Synchronous changes are grouped. |
| `Worksheet#toCSV()` / `Worksheet#loadCSV(text)` / `Worksheet.fromCSV(text)` | CSV serialization. |
| `Workbook#addWorksheet(name?, index?)` / `Workbook#getWorksheetElement(worksheet)` / `Workbook#removeWorksheet(worksheetOrIndex)` | Manage up to 50 sheets. |
| `Workbook#getWorksheetName(worksheet)` / `Workbook#renameWorksheet(worksheet, name)` | Sheet names are owned by the workbook. |
| `Workbook#worksheets` / `Workbook#getWorksheet(nameOrIndex)` | Access sheets. |
| `Workbook#toZip()` / `Workbook.fromZip(data)` | ZIP of CSV files. |
| `cellchange` event | Dispatched by `Worksheet` and re-dispatched by `Workbook` (`CellChangeEvent`). |
| `worksheetadd` / `worksheetremove` / `worksheetrename` events | Dispatched by `Workbook` (`WorksheetEvent`). |

In the browser console, the running app is available as `app.workbook` and `app.grid`.

## Tests

```sh
npm test
```
