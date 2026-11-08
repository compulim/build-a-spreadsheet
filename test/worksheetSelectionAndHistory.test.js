import assert from 'node:assert/strict';
import { test } from 'node:test';

import { Worksheet } from '../src/worksheet/Worksheet.ts';
import { WorksheetSelection } from '../src/worksheet/WorksheetSelection.ts';

const toJSON = selection => JSON.parse(JSON.stringify(selection));

test('Worksheet#selection defaults to A1', () => {
  assert.deepEqual(toJSON(new Worksheet().selection), {
    activeCell: { column: 0, row: 0 },
    range: { left: 0, top: 0, right: 0, bottom: 0 }
  });
});

test('setting Worksheet#selection keeps the instance and dispatches selectionchange', () => {
  const worksheet = new Worksheet();
  const selection = new WorksheetSelection({ column: 1, row: 1 }, { left: 1, top: 1, right: 2, bottom: 3 });
  let count = 0;

  worksheet.addEventListener('selectionchange', () => count++);
  worksheet.selection = selection;

  assert.equal(worksheet.selection, selection);
  assert.equal(count, 1);

  worksheet.selection = new WorksheetSelection({ column: 1, row: 1 }, { left: 1, top: 1, right: 2, bottom: 3 });

  assert.equal(count, 1);
});

test('setting Worksheet#selection does not change the version', () => {
  const worksheet = new Worksheet();
  const { version } = worksheet;

  worksheet.selection = new WorksheetSelection({ column: 3, row: 4 });

  assert.equal(worksheet.version, version);
});

test('Worksheet#selection is clamped and covers merged cells', () => {
  const worksheet = new Worksheet({ columnCount: 5, rowCount: 5 });

  worksheet.edit(edit => edit.mergeCells(1, 1, 2, 2));
  worksheet.selection = new WorksheetSelection({ column: 2, row: 2 }, { left: 2, top: 2, right: 10, bottom: 2 });

  assert.deepEqual(toJSON(worksheet.selection), {
    activeCell: { column: 1, row: 1 },
    range: { left: 1, top: 1, right: 4, bottom: 2 }
  });
});

test('Worksheet#selection keeps an active cell in the middle of the range', () => {
  const worksheet = new Worksheet();

  worksheet.selection = new WorksheetSelection({ column: 1, row: 1 }, { left: 0, top: 0, right: 2, bottom: 2 });

  assert.deepEqual(toJSON(worksheet.selection), {
    activeCell: { column: 1, row: 1 },
    range: { left: 0, top: 0, right: 2, bottom: 2 }
  });
});

test('Worksheet#selection expands when merging cells under it', () => {
  const worksheet = new Worksheet();
  let count = 0;

  worksheet.selection = new WorksheetSelection({ column: 2, row: 2 });
  worksheet.addEventListener('selectionchange', () => count++);
  worksheet.edit(edit => edit.mergeCells(1, 1, 3, 3));

  assert.deepEqual(toJSON(worksheet.selection), {
    activeCell: { column: 1, row: 1 },
    range: { left: 1, top: 1, right: 3, bottom: 3 }
  });
  assert.equal(count, 1);
});

test('Worksheet#undo() and redo() revert, re-apply and select the changed cells', async () => {
  const worksheet = new Worksheet();
  let historyChangeCount = 0;

  worksheet.addEventListener('historychange', () => historyChangeCount++);

  assert.equal(worksheet.canUndo, false);
  assert.equal(worksheet.undo(), false);
  assert.equal(worksheet.redo(), false);

  worksheet.edit(edit => edit.setValues([['a', 'b']], 1, 2));
  await Promise.resolve();

  assert.equal(worksheet.canUndo, true);
  assert.equal(worksheet.canRedo, false);
  assert.equal(historyChangeCount, 1);

  assert.equal(worksheet.undo(), true);
  await Promise.resolve();

  assert.deepEqual(worksheet.getValues(1, 2, 2, 1), [['', '']]);
  assert.equal(worksheet.canUndo, false);
  assert.equal(worksheet.canRedo, true);
  assert.equal(historyChangeCount, 2);
  assert.deepEqual(toJSON(worksheet.selection), {
    activeCell: { column: 1, row: 2 },
    range: { left: 1, top: 2, right: 2, bottom: 2 }
  });

  worksheet.selection = new WorksheetSelection({ column: 0, row: 0 });

  assert.equal(worksheet.redo(), true);
  assert.deepEqual(worksheet.getValues(1, 2, 2, 1), [['a', 'b']]);
  assert.deepEqual(toJSON(worksheet.selection), {
    activeCell: { column: 1, row: 2 },
    range: { left: 1, top: 2, right: 2, bottom: 2 }
  });
});
