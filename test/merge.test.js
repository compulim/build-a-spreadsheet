import assert from 'node:assert/strict';
import { test } from 'node:test';
import { Worksheet } from '../src/workbook/Worksheet.js';

test('mergeCells concatenates values into the first cell and keeps its formatting', () => {
  const worksheet = new Worksheet();

  worksheet.setValues([['a', 1, true]], 1, 1);
  worksheet.setFormattings([[{ bold: true }, { italic: true }]], 1, 1);
  worksheet.mergeCells(1, 1, 3);

  assert.deepEqual(worksheet.getValues(1, 1, 3, 1), [['a 1 TRUE', '', '']]);
  assert.equal(worksheet.getFormattings(1, 1, 1, 1)[0][0].bold, true);
  assert.deepEqual(worksheet.getMerge(2, 1), { x: 1, y: 1, width: 3 });
  assert.deepEqual(worksheet.getMerges(1), [{ x: 1, y: 1, width: 3 }]);
  assert.equal(worksheet.getMerge(4, 1), undefined);
  assert.equal(worksheet.getMerge(1, 2), undefined);
});

test('a single value keeps its type', () => {
  const worksheet = new Worksheet();

  worksheet.setValue(2, 0, 42);
  worksheet.mergeCells(0, 0, 3);

  assert.equal(worksheet.getValue(0, 0), 42);
});

test('unmergeCells keeps content in the first cell', () => {
  const worksheet = new Worksheet();

  worksheet.setValues([['a', 'b']]);
  worksheet.mergeCells(0, 0, 2);
  worksheet.unmergeCells(1, 0);

  assert.equal(worksheet.getMerge(0, 0), undefined);
  assert.deepEqual(worksheet.getValues(0, 0, 2, 1), [['a b', '']]);
});

test('merging removes inner borders and keeps outer borders', () => {
  const worksheet = new Worksheet();
  const range = { x: 1, y: 1, width: 3, height: 1 };

  worksheet.setBorder(range, new Set(['outside', 'inside']));
  worksheet.mergeCells(1, 1, 3);

  assert.deepEqual(worksheet.getBorder(1, 1), { top: true, right: false, bottom: true, left: true });
  assert.deepEqual(worksheet.getBorder(3, 1), { top: true, right: true, bottom: true, left: false });

  worksheet.setBorder(range, new Set(['inside']));
  assert.equal(worksheet.getBorder(2, 1).left, false);
});

test('overlapping merges are combined', () => {
  const worksheet = new Worksheet();

  worksheet.mergeCells(0, 0, 2);
  worksheet.mergeCells(1, 0, 3);

  assert.deepEqual(worksheet.getMerges(0), [{ x: 0, y: 0, width: 4 }]);
});

test('merge is undoable', async () => {
  const worksheet = new Worksheet();

  worksheet.setValues([['a', 'b']]);
  await Promise.resolve();
  worksheet.mergeCells(0, 0, 2);
  worksheet.undo();

  assert.equal(worksheet.getMerge(0, 0), undefined);
  assert.deepEqual(worksheet.getValues(0, 0, 2, 1), [['a', 'b']]);

  worksheet.redo();
  assert.deepEqual(worksheet.getMerge(1, 0), { x: 0, y: 0, width: 2 });
});
