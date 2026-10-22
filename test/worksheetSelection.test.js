import assert from 'node:assert/strict';
import { test } from 'node:test';

import { WorksheetSelection } from '../src/WorksheetSelection.js';

test('WorksheetSelection defaults the range to the active cell', () => {
  const selection = new WorksheetSelection({ column: 2, row: 3 });

  assert.deepEqual(selection.activeCell, { column: 2, row: 3 });
  assert.deepEqual(selection.range, { left: 2, top: 3, right: 2, bottom: 3 });
  assert.deepEqual(selection.focus, { column: 2, row: 3 });
});

test('WorksheetSelection focus is the corner opposite to the active cell', () => {
  const range = { left: 1, top: 1, right: 4, bottom: 5 };

  assert.deepEqual(new WorksheetSelection({ column: 1, row: 1 }, range).focus, { column: 4, row: 5 });
  assert.deepEqual(new WorksheetSelection({ column: 4, row: 5 }, range).focus, { column: 1, row: 1 });
  assert.deepEqual(new WorksheetSelection({ column: 4, row: 1 }, range).focus, { column: 1, row: 5 });
});

test('WorksheetSelection is immutable and must contain the active cell', () => {
  const selection = new WorksheetSelection({ column: 0, row: 0 }, { left: 0, top: 0, right: 1, bottom: 1 });

  assert.throws(() => (selection.activeCell.column = 1), TypeError);
  assert.throws(() => (selection.range.left = 1), TypeError);
  assert.throws(() => new WorksheetSelection({ column: 5, row: 0 }, { left: 0, top: 0, right: 1, bottom: 1 }), RangeError);
});
