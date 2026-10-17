import assert from 'node:assert/strict';
import { describe, test } from 'node:test';

import { CellChangeEvent } from '../src/workbook/private/CellChangeEvent.js';
import { COLUMN_COUNT, MAX_WORKSHEETS, ROW_COUNT } from '../src/workbook/constants.js';
import { getColumnName } from '../src/workbook/getColumnName.js';
import { Workbook } from '../src/workbook/Workbook.js';
import { Worksheet } from '../src/workbook/Worksheet.js';
import { createZip } from '../src/zip.js';

describe('Worksheet', () => {
  test('dimensions', () => {
    const worksheet = new Worksheet();

    assert.equal(worksheet.columnCount, COLUMN_COUNT);
    assert.equal(worksheet.rowCount, ROW_COUNT);
    assert.equal(COLUMN_COUNT, 50);
    assert.equal(ROW_COUNT, 10_000);
  });

  test('setValue and getValues', () => {
    const worksheet = new Worksheet();

    assert.deepEqual(worksheet.getValues(0, 0, 0, 0), []);

    worksheet.setValue(0, 0, 'A1');
    worksheet.setValue(2, 1, 3);
    worksheet.setValue(1, 3, true);

    assert.equal(worksheet.getValue(2, 1), 3);
    assert.equal(worksheet.getValue(49, 9999), undefined);
    assert.deepEqual(worksheet.getValues(0, 0, 3, 4), [
      ['A1', '', ''],
      ['', '', 3],
      ['', '', ''],
      ['', true, '']
    ]);

    worksheet.setValue(1, 3, '');
    worksheet.setValue(2, 1, null);

    assert.deepEqual(worksheet.getValues(0, 0, 2, 2), [
      ['A1', ''],
      ['', '']
    ]);
    assert.deepEqual(worksheet.getValues(2, 1, 1, 1), [['']]);
    assert.throws(() => worksheet.getValues(49, 0, 2, 1), RangeError);
  });

  test('setValues with empty values clears cells', () => {
    const worksheet = new Worksheet();

    worksheet.setValues([[1, 2], [3]]);
    worksheet.setValues([['', null]], 0, 0);

    assert.deepEqual(worksheet.getValues(0, 0, 2, 2), [
      ['', ''],
      [3, '']
    ]);
  });

  test('strings are trimmed and blank strings clear the cell', () => {
    const worksheet = new Worksheet();

    worksheet.setValue(0, 0, '  hello \n');
    assert.equal(worksheet.getValue(0, 0), 'hello');

    worksheet.setValues([['   ', ' x ']]);
    assert.equal(worksheet.getValue(0, 0), undefined);
    assert.equal(worksheet.getValue(1, 0), 'x');

    worksheet.setValue(1, 0, () => ' \t ');
    assert.equal(worksheet.getValue(1, 0), undefined);
    assert.equal(worksheet.toCSV(), '');
  });

  test('updater functions', () => {
    const worksheet = new Worksheet();
    const calls = [];

    worksheet.setValue(0, 0, 1);
    worksheet.setValue(0, 0, (value, x, y) => {
      calls.push([value, x, y]);

      return value + 1;
    });
    worksheet.setValue(1, 0, value => (value === undefined ? 'empty' : value));

    assert.deepEqual(calls, [[1, 0, 0]]);
    assert.deepEqual(worksheet.getValues(0, 0, 2, 1), [[2, 'empty']]);

    worksheet.setValues((_, x, y) => x * 10 + y, 1, 1, 2, 2);
    assert.deepEqual(worksheet.getValues(1, 1, 2, 2), [
      [11, 21],
      [12, 22]
    ]);

    worksheet.setValues(() => '', 0, 0, 50, 10_000);
    assert.equal(worksheet.toCSV(), '');
    assert.throws(() => worksheet.setValues(() => 1, 0, 0), RangeError);
  });

  test('column widths', () => {
    const worksheet = new Worksheet();
    const events = [];

    worksheet.addEventListener('columnwidthchange', event => events.push(event.x));

    assert.equal(worksheet.getColumnWidth(0), 100);

    worksheet.setColumnWidth(2, 150);
    worksheet.setColumnWidth(2, 150);

    assert.equal(worksheet.getColumnWidth(2), 150);
    assert.deepEqual(events, [2]);
    assert.throws(() => worksheet.setColumnWidth(50, 10), RangeError);
  });

  test('undo and redo', async () => {
    const tick = () => new Promise(resolve => setTimeout(resolve));
    const worksheet = new Worksheet();

    assert.equal(worksheet.undo(), false);
    assert.equal(worksheet.redo(), false);

    worksheet.setValue(0, 0, 'a');
    await tick();
    worksheet.setValues([['b', 'c']]);
    await tick();

    assert.equal(worksheet.canUndo, true);
    assert.equal(worksheet.undo(), true);
    assert.deepEqual(worksheet.getValues(0, 0, 2, 1), [['a', '']]);
    assert.equal(worksheet.undo(), true);
    assert.deepEqual(worksheet.getValues(0, 0, 2, 1), [['', '']]);
    assert.equal(worksheet.undo(), false);

    assert.equal(worksheet.redo(), true);
    assert.equal(worksheet.redo(), true);
    assert.equal(worksheet.redo(), false);
    assert.deepEqual(worksheet.getValues(0, 0, 2, 1), [['b', 'c']]);

    worksheet.undo();
    assert.equal(worksheet.canRedo, true);
    worksheet.setValue(1, 0, 2);
    assert.equal(worksheet.canRedo, false);
  });

  test('setValue validates input', () => {
    const worksheet = new Worksheet();

    assert.throws(() => worksheet.setValue(50, 0, 1), RangeError);
    assert.throws(() => worksheet.setValue(0, 10_000, 1), RangeError);
    assert.throws(() => worksheet.setValue(-1, 0, 1), RangeError);
    assert.throws(() => worksheet.setValue(0.5, 0, 1), RangeError);
    assert.throws(() => worksheet.setValue(0, 0, {}), TypeError);
    assert.throws(() => worksheet.setValue(0, 0, NaN), TypeError);

    worksheet.setValue(49, 9999, 'last');
    assert.equal(worksheet.getValue(49, 9999), 'last');
  });

  test('cellchange event', () => {
    const worksheet = new Worksheet();
    const events = [];

    worksheet.addEventListener('cellchange', event => events.push(event));

    worksheet.setValue(1, 2, 'x');
    worksheet.setValue(1, 2, 'x'); // No change, no event.
    worksheet.setValue(1, 2, 1);

    assert.equal(events.length, 2);
    assert.ok(events[0] instanceof CellChangeEvent);
    assert.equal(events[0].type, 'cellchange');
    assert.deepEqual(
      events.map(({ x, y, value, previousValue, worksheet }) => ({
        x,
        y,
        value,
        previousValue,
        worksheet
      })),
      [
        { x: 1, y: 2, value: 'x', previousValue: '', worksheet },
        { x: 1, y: 2, value: 1, previousValue: 'x', worksheet }
      ]
    );
  });

  test('CSV round trip', () => {
    const worksheet = new Worksheet();

    worksheet.setValue(0, 0, 'Name, "quoted"');
    worksheet.setValue(1, 0, 42);
    worksheet.setValue(2, 1, false);

    const csv = worksheet.toCSV();

    assert.equal(csv, '"Name, ""quoted""",42,\r\n,,FALSE\r\n');
    assert.deepEqual(Worksheet.fromCSV(csv).getValues(0, 0, 6, 6), worksheet.getValues(0, 0, 6, 6));
    assert.equal(Worksheet.fromCSV(csv).canUndo, false);

  });

  test('CSV larger than worksheet is clipped', () => {
    const csv = Array.from({ length: 10_005 }, () => new Array(55).fill(1).join(',')).join('\n');
    const worksheet = Worksheet.fromCSV(csv);

    assert.equal(worksheet.getValue(49, 9999), 1);
    assert.equal(worksheet.toCSV().split('\r\n').length - 1, 10_000);
  });

});

describe('Workbook', () => {
  test('is an EventTarget re-dispatching cellchange', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();
    const events = [];

    assert.ok(workbook instanceof EventTarget);
    assert.equal(workbook.getWorksheetName(worksheet), 'Sheet1');

    workbook.addEventListener('cellchange', event => events.push(event));
    worksheet.setValue(0, 0, 'hello');

    assert.equal(events.length, 1);
    assert.equal(events[0].worksheet, worksheet);
    assert.equal(events[0].value, 'hello');

    workbook.removeWorksheet(worksheet);
    worksheet.setValue(0, 0, 'bye');

    assert.equal(events.length, 1);
  });

  test('up to 50 worksheets with unique names', () => {
    const workbook = new Workbook();

    for (let index = 0; index < MAX_WORKSHEETS; index++) {
      workbook.addWorksheet();
    }

    assert.equal(workbook.worksheets.length, 50);
    assert.equal(new Set(workbook.worksheets.map(worksheet => workbook.getWorksheetName(worksheet))).size, 50);
    assert.throws(() => workbook.addWorksheet(), RangeError);

    workbook.removeWorksheet(0);

    assert.throws(() => workbook.addWorksheet('sheet2'), TypeError);
  });

  test('rename', () => {
    const workbook = new Workbook();
    const sheet1 = workbook.addWorksheet();
    const sheet2 = workbook.addWorksheet();
    const events = [];

    workbook.addEventListener('worksheetrename', event => events.push(event));

    assert.throws(() => workbook.renameWorksheet(sheet2, 'SHEET1'), TypeError);
    assert.throws(() => workbook.renameWorksheet(sheet2, 'a/b'), TypeError);
    assert.throws(() => workbook.addWorksheet(''), TypeError);
    assert.throws(() => workbook.addWorksheet('x'.repeat(32)), TypeError);

    workbook.renameWorksheet(sheet1, 'Data');

    assert.equal(workbook.getWorksheet('data'), sheet1);
    assert.equal(events.length, 1);
    assert.equal(events[0].previousName, 'Sheet1');
  });

  test('ZIP round trip', async () => {
    const workbook = new Workbook();
    const first = workbook.addWorksheet('First');
    const second = workbook.addWorksheet('第二');

    workbook.addWorksheet('Empty');
    first.setValue(0, 0, 'a');
    first.setValue(3, 2, 1.25);
    second.setValue(49, 9999, true);

    const zip = await workbook.toZip();
    const loaded = await Workbook.fromZip(zip);

    assert.deepEqual(
      loaded.worksheets.map(worksheet => loaded.getWorksheetName(worksheet)),
      ['First', '第二', 'Empty']
    );
    assert.equal(loaded.worksheets[0].toCSV(), first.toCSV());
    assert.equal(loaded.worksheets[1].getValue(49, 9999), true);
    assert.equal(loaded.worksheets[2].toCSV(), '');
  });

  test('fromZip ignores non-CSV files and folders', async () => {
    const encoder = new TextEncoder();
    const zip = await createZip([
      { name: 'folder/Sales.CSV', data: encoder.encode('1,2\n') },
      { name: 'readme.txt', data: encoder.encode('hi') },
      { name: '__MACOSX/folder/._Sales.CSV', data: encoder.encode('junk') }
    ]);

    const workbook = await Workbook.fromZip(zip);

    assert.deepEqual(
      workbook.worksheets.map(worksheet => workbook.getWorksheetName(worksheet)),
      ['Sales']
    );
    assert.deepEqual(workbook.worksheets[0].getValues(0, 0, 3, 1), [[1, 2, '']]);
  });
});

test('getColumnName', () => {
  assert.equal(getColumnName(0), 'A');
  assert.equal(getColumnName(25), 'Z');
  assert.equal(getColumnName(26), 'AA');
  assert.equal(getColumnName(49), 'AX');
  assert.equal(getColumnName(701), 'ZZ');
  assert.equal(getColumnName(702), 'AAA');
});

describe('Worksheet borders', () => {
  const range = { x: 1, y: 1, width: 2, height: 2 };
  const none = { top: false, right: false, bottom: false, left: false };

  for (const side of /** @type {const} */ (['left', 'right', 'top', 'bottom'])) {
    test(side, () => {
      const worksheet = new Worksheet();

      worksheet.setBorder(range, new Set([side]));

      for (let y = 1; y <= 2; y++) {
        for (let x = 1; x <= 2; x++) {
          const expected =
            (side === 'left' && x === 1) ||
            (side === 'right' && x === 2) ||
            (side === 'top' && y === 1) ||
            (side === 'bottom' && y === 2);

          assert.deepEqual(worksheet.getBorder(x, y), { ...none, [side]: expected });
        }
      }
    });
  }

  test('outside, inside and unset', () => {
    const worksheet = new Worksheet();
    let count = 0;

    worksheet.addEventListener('borderchange', () => count++);
    worksheet.setBorder(range, new Set(['outside']));

    assert.deepEqual(worksheet.getBorder(1, 1), { ...none, top: true, left: true });
    assert.deepEqual(worksheet.getBorder(2, 2), { ...none, bottom: true, right: true });
    // One event per stored edge (left/top) changed: 4 on the left/top edges, 4 on the right/bottom neighbors.
    assert.equal(count, 8);

    worksheet.setBorder(range, new Set(['inside']));
    assert.deepEqual(worksheet.getBorder(1, 2), { top: true, right: true, bottom: true, left: true });

    worksheet.setBorder(range, new Set(['outside']), false);
    assert.deepEqual(worksheet.getBorder(1, 1), { ...none, right: true, bottom: true });
    assert.deepEqual(worksheet.getBorder(0, 1), none);

    worksheet.setBorder(range, new Set(['inside']), false);
    assert.deepEqual(worksheet.getBorder(1, 2), none);
  });

  test('a Set of edges is applied in a single call', () => {
    const worksheet = new Worksheet();

    worksheet.setBorder(range, new Set(['top', 'left']));
    assert.deepEqual(worksheet.getBorder(1, 1), { ...none, top: true, left: true });
    assert.deepEqual(worksheet.getBorder(2, 2), none);

    worksheet.setBorder(range, new Set(['outside', 'inside']));
    assert.deepEqual(worksheet.getBorder(1, 1), { top: true, right: true, bottom: true, left: true });
    assert.deepEqual(worksheet.getBorder(2, 2), { top: true, right: true, bottom: true, left: true });
  });

  test('whole column and whole row', () => {
    const worksheet = new Worksheet();

    worksheet.setBorder({ x: 2 }, new Set(['left']));
    worksheet.setBorder({ x: 2 }, new Set(['right']));
    assert.deepEqual(worksheet.getBorder(2, 0), { ...none, left: true, right: true });
    assert.deepEqual(worksheet.getBorder(2, worksheet.rowCount - 1), { ...none, left: true, right: true });

    worksheet.setBorder({ y: 4 }, new Set(['top']));
    worksheet.setBorder({ y: 4 }, new Set(['bottom']));
    assert.deepEqual(worksheet.getBorder(0, 4), { ...none, top: true, bottom: true });
    assert.deepEqual(worksheet.getBorder(worksheet.columnCount - 1, 4), { ...none, top: true, bottom: true });
  });

  test('right and bottom are shared with the next cell, and skipped on the last column and row', () => {
    const worksheet = new Worksheet();
    const lastX = worksheet.columnCount - 1;
    const lastY = worksheet.rowCount - 1;

    worksheet.setBorder({ x: 0, y: 0 }, new Set(['right', 'bottom']));
    assert.deepEqual(worksheet.getBorder(1, 0), { ...none, left: true });
    assert.deepEqual(worksheet.getBorder(0, 1), { ...none, top: true });

    worksheet.setBorder({ x: 1, y: 0 }, new Set(['left']), false);
    assert.deepEqual(worksheet.getBorder(0, 0), { ...none, bottom: true });

    worksheet.setBorder({ x: lastX, y: lastY }, new Set(['right', 'bottom']));
    assert.deepEqual(worksheet.getBorder(lastX, lastY), none);
  });
});

test('Worksheet formatting', () => {
  const worksheet = new Worksheet({ columnCount: 3, rowCount: 3 });
  const events = [];

  worksheet.setValues([
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9]
  ]);
  worksheet.addEventListener('formattingchange', ({ x, y }) => events.push([x, y]));

  assert.deepStrictEqual(worksheet.getFormattings(0, 0, 1, 1)[0][0], {
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    align: undefined,
    highlight: undefined,
    format: undefined
  });

  worksheet.setFormattings([[{ align: 'left' }]], 0, 2);
  assert.strictEqual(worksheet.getFormattings(0, 2, 1, 1)[0][0].align, 'left');
  worksheet.setFormattings([[null]], 0, 2);
  events.length = 0;

  worksheet.setFormattings([[{ bold: true, underline: true, align: 'center' }]], 1, 1);
  assert.deepStrictEqual(worksheet.getFormattings(1, 1, 1, 1)[0][0], {
    bold: true,
    italic: false,
    underline: true,
    strikethrough: false,
    align: 'center',
    highlight: undefined,
    format: undefined
  });
  assert.deepStrictEqual(events, [[1, 1]]);

  worksheet.setFormattings(formatting => ({ ...formatting, bold: true }), 1, 1, 1, 1);
  assert.strictEqual(events.length, 1);

  worksheet.setFormattings(formatting => ({ ...formatting, italic: true }), 0, 0, 2, 2);
  assert.deepStrictEqual(
    worksheet.getFormattings(0, 0, 2, 2).map(row => row.map(({ italic, align }) => [italic, align])),
    [
      [[true, undefined], [true, undefined]],
      [[true, undefined], [true, 'center']]
    ]
  );

  worksheet.setFormattings([[null, { strikethrough: true, align: 'right' }]], 1, 1);
  assert.deepStrictEqual(worksheet.getFormattings(1, 1, 1, 1)[0][0].bold, false);
  assert.deepStrictEqual(worksheet.getFormattings(2, 1, 1, 1)[0][0].align, 'right');

  assert.throws(() => worksheet.setFormattings([[{ align: 'justify' }]], 0, 0), TypeError);

  worksheet.setFormattings([[{ highlight: 'green' }]], 0, 0);
  assert.strictEqual(worksheet.getFormattings(0, 0, 1, 1)[0][0].highlight, 'green');
  worksheet.setFormattings(formatting => ({ ...formatting, highlight: undefined }), 0, 0, 1, 1);
  assert.strictEqual(worksheet.getFormattings(0, 0, 1, 1)[0][0].highlight, undefined);
  assert.throws(() => worksheet.setFormattings([[{ highlight: 1 }]], 0, 0), TypeError);

  worksheet.setFormattings([[{ format: 'currency' }]], 0, 0);
  assert.strictEqual(worksheet.getFormattings(0, 0, 1, 1)[0][0].format, 'currency');
  worksheet.setFormattings(formatting => ({ ...formatting, format: undefined }), 0, 0, 1, 1);
  assert.strictEqual(worksheet.getFormattings(0, 0, 1, 1)[0][0].format, undefined);
  assert.throws(() => worksheet.setFormattings([[{ format: 'scientific' }]], 0, 0), TypeError);
});

test('Worksheet formatting is not kept on empty cells', () => {
  const worksheet = new Worksheet({ columnCount: 3, rowCount: 3 });

  worksheet.setValue(0, 0, 'a');
  worksheet.setFormattings(() => ({ bold: true }), 0, 0, 2, 1);
  assert.strictEqual(worksheet.getFormattings(0, 0, 1, 1)[0][0].bold, true);
  assert.strictEqual(worksheet.getFormattings(1, 0, 1, 1)[0][0].bold, false);

  const events = [];

  worksheet.addEventListener('formattingchange', ({ x, y }) => events.push([x, y]));
  worksheet.setValue(0, 0, '');
  assert.deepStrictEqual(events, [[0, 0]]);

  worksheet.setValue(0, 0, 'b');
  assert.strictEqual(worksheet.getFormattings(0, 0, 1, 1)[0][0].bold, false);
});

test('Worksheet undo/redo formatting', async () => {
  const nextTask = () => new Promise(resolve => setTimeout(resolve));
  const worksheet = new Worksheet({ columnCount: 3, rowCount: 3 });

  worksheet.setValue(0, 0, 'ABC');
  await nextTask();
  worksheet.setFormattings([[{ bold: true }]], 0, 0);
  await nextTask();

  assert.strictEqual(worksheet.undo(), true);
  assert.strictEqual(worksheet.getValue(0, 0), 'ABC');
  assert.strictEqual(worksheet.getFormattings(0, 0, 1, 1)[0][0].bold, false);

  assert.strictEqual(worksheet.undo(), true);
  assert.strictEqual(worksheet.getValue(0, 0), undefined);

  worksheet.redo();
  worksheet.redo();
  assert.strictEqual(worksheet.getFormattings(0, 0, 1, 1)[0][0].bold, true);
  await nextTask();

  // Deleting a formatted value, then undo, restores both value and formatting.
  worksheet.setValue(0, 0, '');
  await nextTask();
  assert.strictEqual(worksheet.getFormattings(0, 0, 1, 1)[0][0].bold, false);
  worksheet.undo();
  assert.strictEqual(worksheet.getValue(0, 0), 'ABC');
  assert.strictEqual(worksheet.getFormattings(0, 0, 1, 1)[0][0].bold, true);
  worksheet.redo();
  assert.strictEqual(worksheet.getValue(0, 0), undefined);
  assert.strictEqual(worksheet.getFormattings(0, 0, 1, 1)[0][0].bold, false);
});
