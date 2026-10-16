import assert from 'node:assert/strict';
import { test } from 'node:test';

import { UndoManager } from '../src/UndoManager.js';
import { Workbook } from '../src/workbook/Workbook.js';

const tick = () => new Promise(resolve => setTimeout(resolve));

test('groups synchronous changes and undoes/redoes them', async () => {
  const workbook = new Workbook();
  const worksheet = workbook.addWorksheet();
  const undoManager = new UndoManager();

  undoManager.track(workbook);
  worksheet.setValue(0, 0, 'a');
  await tick();
  worksheet.setValues([['b', 'c']]);
  await tick();

  assert.equal(undoManager.canUndo, true);
  assert.equal(undoManager.undo()?.length, 2);
  assert.deepEqual(worksheet.getValues(), [['a']]);
  undoManager.undo();
  assert.deepEqual(worksheet.getValues(), []);
  assert.equal(undoManager.canUndo, false);

  undoManager.redo();
  undoManager.redo();
  assert.deepEqual(worksheet.getValues(), [['b', 'c']]);
  assert.equal(undoManager.canRedo, false);
});

test('new change clears redo', async () => {
  const workbook = new Workbook();
  const worksheet = workbook.addWorksheet();
  const undoManager = new UndoManager();

  undoManager.track(workbook);
  worksheet.setValue(0, 0, 1);
  await tick();
  undoManager.undo();
  assert.equal(undoManager.canRedo, true);
  worksheet.setValue(1, 0, 2);
  assert.equal(undoManager.canRedo, false);
  assert.equal(undoManager.redo(), undefined);
});
