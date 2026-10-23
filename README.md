# build-a-spreadsheet

A high-performance spreadsheet web app that runs entirely in the browser — no server required.
The grid is written in plain JavaScript (ES modules). The page layout and toolbar use [GitHub Primer](https://primer.style/) React components. Everything is bundled with tsup into static files that can be served from GitHub Pages or any CDN.

## Features

- 50 columns × 10,000 rows per sheet, with column (A, B, … AX) and row headers.
- Resize a column by dragging the right border of its header; double-click the border to auto-fit the column to its content. When multiple whole columns are selected, all of them are resized to the same width.
- Virtualized rendering: only the visible rows are in the DOM, while scrolling uses the browser's native scrollbars.
- Smooth scrolling: rows are rendered synchronously on scroll with a viewport of pre-rendered rows on each side, so wheel, touchpad and momentum scrolling stay seamless. Keyboard navigation scrolls smoothly for short distances (instant for long jumps, or when the user prefers reduced motion).
- Keyboard navigation: arrow keys, <kbd>Ctrl</kbd>+arrow (jump to edge), <kbd>Home</kbd>/<kbd>End</kbd>, <kbd>Page Up</kbd>/<kbd>Page Down</kbd>, <kbd>Enter</kbd>/<kbd>Tab</kbd>, <kbd>Alt</kbd>+arrow (move the active cell within the selection), <kbd>Escape</kbd> (collapse the selection to the active cell).
- Multiple selection with <kbd>Shift</kbd>+click, <kbd>Shift</kbd>+arrow keys, mouse drag, or by clicking row/column headers; <kbd>Ctrl</kbd>+<kbd>A</kbd> selects all.
- Editing: type to replace, <kbd>F2</kbd> or double-click to edit, <kbd>Escape</kbd> to cancel, <kbd>Delete</kbd> to clear. A leading `'` keeps the value as text.
- Copy, cut and paste with Excel and Excel Web App (tab-separated text and HTML table on the clipboard). Pasting a single value into a larger selection fills it.
- Values are booleans, numbers or strings. Formulas are not supported.
- Up to 50 named sheets in the bottom bar: <kbd>+</kbd> to add, double-click to rename, × to delete.
- Toolbar built with Primer: `ActionBar` for file, undo/redo, bold/italic/underline/strikethrough, border, merge and unformat; `SegmentedControl` for alignment (Auto/Left/Center/Right) and the highlight marker (None or a colour); `ActionMenu` for the number format; `PageLayout` for the overall page.
- Lock switch: makes the rows/columns listed in `LOCKED_ROWS` and `LOCKED_COLUMNS` in [`src/app.js`](src/app.js) read-only.
- Open a CSV file (added as a new sheet) or a ZIP file of CSV files (replaces the workbook). Save the current sheet as CSV, or the whole workbook as a ZIP of CSV files.

## Running locally

```sh
npm install
npm start # Builds on change and serves ./public/
```

To deploy, run `npm run build` and publish the `public/` folder. All paths are relative.

## Programming interface

```js
import { Workbook } from './src/workbook/Workbook.js';
import { Worksheet } from './src/worksheet/Worksheet.js';

const workbook = new Workbook(); // EventTarget
const sheet = workbook.addWorksheet('Data'); // Worksheet, also an EventTarget

// One `change` event per transaction, with all changes in order.
workbook.addEventListener('change', ({ worksheet, changes }) => {
  for (const change of changes) {
    change.type === 'value' &&
      console.log(workbook.getWorksheetName(worksheet), change.x, change.y, change.previousValue, '→', change.value);
  }
});

// All modifications are made through the `edit` object, which is only usable inside the callback.
sheet.edit(edit => {
  edit.setValue(0, 0, 'Name'); // setValue(x: number, y: number, value: boolean | number | string), zero-based
  edit.setValue(1, 0, 42);
  edit.setValue(2, 0, true);
});
sheet.getValues(0, 0, 3, 1); // [['Name', 42, true]] — getValues(x, y, width, height), empty cells are ''
sheet.edit(edit => edit.setValues([['', '']], 1, 0)); // Clears B1:C1
sheet.setColumnWidth(0, 150); // Outside of edit(), not undoable

sheet.toCSV(); // 'Name,42,TRUE\r\n'
workbook.addWorksheet('Imported').edit(edit => edit.loadCSV('a,1\r\n'));

const zip = await workbook.toZip(); // Uint8Array, one "<sheet name>.csv" per worksheet
const loaded = await Workbook.fromZip(zip);
```

In a browser where `<compulim-worksheet>` ([`src/worksheet/WorksheetElement.js`](src/worksheet/WorksheetElement.js)) is defined, the workbook instantiates one element per worksheet. The element renders the grid, handles user interactions, and wholly owns the worksheet. Its public members are the read-only `worksheet` property, `selection`, `undo()`, `redo()`, `canUndo` and `canRedo`:

```js
import './src/worksheet/WorksheetElement.js';

const worksheet = workbook.addWorksheet('Hello');
const element = workbook.getWorksheetElement(worksheet); // <compulim-worksheet>, element.worksheet === worksheet

element.setAttribute('locked-row-count', '1'); // Number of lockable rows, can be changed at any time
element.setAttribute('locked-column-count', '1'); // Number of lockable columns
document.body.append(element);
element.toggleAttribute('locked', true); // Locks row 1
worksheet.edit(edit => edit.setValue(0, 1, 'Hello')); // The grid updates
```

The element keeps the undo history: it subscribes to the change events of its worksheet and remembers the `undo()` object of each `edit()` transaction (and the `redo()` object returned by undoing it) in a stack. The history logic lives in [`src/worksheet/WorksheetHistory.js`](src/worksheet/WorksheetHistory.js), which has no DOM dependency. The selection is an immutable [`WorksheetSelection`](src/worksheet/WorksheetSelection.js):

```js
import { WorksheetSelection } from './src/worksheet/WorksheetSelection.js';

element.undo(); // true, reverted and selected the changed cells; false if nothing to undo
element.redo();
element.canUndo; // Re-check `canUndo` / `canRedo` on the element's `historychange` event

element.selection; // WorksheetSelection { activeCell: { column, row }, range: { left, top, right, bottom } }
element.selection = new WorksheetSelection({ column: 0, row: 0 }, { left: 0, top: 0, right: 2, bottom: 1 }); // Selects A1:C2
```

| Member | Description |
| --- | --- |
| `Worksheet#edit(callback)` | Runs `callback(edit)` as a transaction and returns `{ undo() }`. Each transaction with changes sets a new `version` (GUID); `undo()` throws unless the worksheet is at the version right after the transaction, restores the previous version and returns `{ redo() }`, which likewise requires that previous version. `edit` has all editing functions below and is only usable inside the callback. A nested `edit()` joins the outer transaction. If the callback throws, all its changes are rolled back. Otherwise, a single `change` event is dispatched if anything changed. |
| `edit.setValue(x, y, valueOrUpdater)` | Sets a cell (zero-based), or updates it with `(value, x, y) => nextValue`. Strings are trimmed. Blank strings, `null` or `undefined` clear it. Throws `RangeError` / `TypeError` on invalid input. |
| `edit.setValues(values, x?, y?)` / `edit.setValues(updater, x, y, width, height)` | Sets a block of values (use `''` to clear), or updates every cell of a range with `(value, x, y) => nextValue`. Values outside of the sheet are ignored. |
| `edit.setFormattings(...)` / `edit.setBorder(range, edges, value?)` / `edit.mergeCells(x, y, width, height?)` / `edit.unmergeCells(x, y)` | Formatting, borders and merged cells. |
| `edit.loadCSV(text)` | Loads CSV, the `change` event has `load: true`, which clears undo history. |
| `Worksheet#getValue(x, y)` | Returns a cell value, `undefined` if empty. |
| `Worksheet#getValues(x, y, width, height)` | Returns a 2D array (`values[y][x]`) of a range. |
| `Worksheet#getColumnWidth(x)` | Column width in pixels. |
| `Worksheet#setColumnWidth(x, width)` | Sets a column width in pixels, outside of `edit()`: not undoable and does not change `version`. Dispatches `columnwidthchange` (`CustomEvent`, `detail` of `{ x, width, previousWidth }`) if changed. |
| `Worksheet#toCSV()` / `Worksheet.fromCSV(text)` | CSV serialization. |
| `Workbook#addWorksheet(name?, index?)` / `Workbook#getWorksheetElement(worksheet)` / `Workbook#removeWorksheet(worksheetOrIndex)` | Manage up to 50 sheets. |
| `Workbook#getWorksheetName(worksheet)` / `Workbook#renameWorksheet(worksheet, name)` | Sheet names are owned by the workbook. |
| `Workbook#worksheets` / `Workbook#getWorksheet(nameOrIndex)` | Access sheets. |
| `Workbook#toZip()` / `Workbook.fromZip(data)` | ZIP of CSV files. |
| `change` event | Dispatched by `Worksheet` once per transaction and re-dispatched by `Workbook` (`WorksheetChangeEvent`, with `worksheet`, `load` and `changes`). Each change has a `type` (`'value'`, `'formatting'`, `'border'` or `'merge'`) with the new and previous value. |
| `worksheetadd` / `worksheetremove` / `worksheetrename` events | Dispatched by `Workbook` (`WorksheetEvent`). |

In the browser console, the running app is available as `app.workbook` and `app.grid`.

## Tests

```sh
npm test
```
