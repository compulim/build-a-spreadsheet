import assert from "node:assert/strict";
import { test } from "node:test";
import { Worksheet } from "../src/worksheet/Worksheet.ts";
import { WorksheetHistory } from "../src/worksheet/WorksheetHistory.ts";

test("mergeCells concatenates values into the first cell and keeps its formatting", () => {
  const worksheet = new Worksheet();

  worksheet.edit((edit) => edit.setValues([["a", 1, true]], 1, 1));
  worksheet.edit((edit) =>
    edit.setFormattings([[{ bold: true }, { italic: true }]], 1, 1),
  );
  worksheet.edit((edit) => edit.mergeCells(1, 1, 3));

  assert.deepEqual(worksheet.getValues(1, 1, 3, 1), [["a 1 TRUE", "", ""]]);
  assert.equal(worksheet.getFormattings(1, 1, 1, 1)[0][0].bold, true);
  assert.deepEqual(worksheet.getMerges(2, 1, 1, 1), [
    { x: 1, y: 1, width: 3, height: 1 },
  ]);
  assert.deepEqual(worksheet.getMerges(), [
    { x: 1, y: 1, width: 3, height: 1 },
  ]);
  assert.deepEqual(worksheet.getMerges(4, 1, 1, 1), []);
  assert.deepEqual(worksheet.getMerges(1, 2, 1, 1), []);
});

test("a single value keeps its type", () => {
  const worksheet = new Worksheet();

  worksheet.edit((edit) => edit.setValue(2, 0, 42));
  worksheet.edit((edit) => edit.mergeCells(0, 0, 3));

  assert.equal(worksheet.getValue(0, 0), 42);
});

test("unmergeCells keeps content in the first cell", () => {
  const worksheet = new Worksheet();

  worksheet.edit((edit) => edit.setValues([["a", "b"]]));
  worksheet.edit((edit) => edit.mergeCells(0, 0, 2));
  worksheet.edit((edit) => edit.unmergeCells(1, 0));

  assert.deepEqual(worksheet.getMerges(0, 0, 1, 1), []);
  assert.deepEqual(worksheet.getValues(0, 0, 2, 1), [["a b", ""]]);
});

test("merging removes inner borders and keeps outer borders", () => {
  const worksheet = new Worksheet();
  const range = { x: 1, y: 1, width: 3, height: 1 };

  worksheet.edit((edit) =>
    edit.setBorder(range, new Set(["outside", "inside"])),
  );
  worksheet.edit((edit) => edit.mergeCells(1, 1, 3));

  assert.deepEqual(worksheet.getBorder(1, 1), {
    top: true,
    right: false,
    bottom: true,
    left: true,
  });
  assert.deepEqual(worksheet.getBorder(3, 1), {
    top: true,
    right: true,
    bottom: true,
    left: false,
  });

  worksheet.edit((edit) => edit.setBorder(range, new Set(["inside"])));
  assert.equal(worksheet.getBorder(2, 1).left, false);
});

test("overlapping merges are combined", () => {
  const worksheet = new Worksheet();

  worksheet.edit((edit) => edit.mergeCells(0, 0, 2));
  worksheet.edit((edit) => edit.mergeCells(1, 0, 3));

  assert.deepEqual(worksheet.getMerges(), [
    { x: 0, y: 0, width: 4, height: 1 },
  ]);
});

test("getMerges returns all merges", () => {
  const worksheet = new Worksheet();

  worksheet.edit((edit) => edit.mergeCells(1, 1, 2));
  worksheet.edit((edit) => edit.mergeCells(2, 3, 3));

  assert.deepEqual(
    worksheet.getMerges().sort((a, b) => a.y - b.y),
    [
      { x: 1, y: 1, width: 2, height: 1 },
      { x: 2, y: 3, width: 3, height: 1 },
    ],
  );
});

test("merge is undoable", async () => {
  const worksheet = new Worksheet();
  const history = new WorksheetHistory(worksheet);

  worksheet.edit((edit) => edit.setValues([["a", "b"]]));
  await Promise.resolve();
  worksheet.edit((edit) => edit.mergeCells(0, 0, 2));
  history.undo();

  assert.deepEqual(worksheet.getMerges(0, 0, 1, 1), []);
  assert.deepEqual(worksheet.getValues(0, 0, 2, 1), [["a", "b"]]);

  history.redo();
  assert.deepEqual(worksheet.getMerges(1, 0, 1, 1), [
    { x: 0, y: 0, width: 2, height: 1 },
  ]);
});

test("merge uses the formatting of the first non-empty cell", () => {
  const worksheet = new Worksheet();

  worksheet.edit((edit) => edit.setValues([["", "b", "c"]], 1, 1));
  worksheet.edit((edit) =>
    edit.setFormattings([[null, { italic: true }, { bold: true }]], 1, 1),
  );
  worksheet.edit((edit) => edit.mergeCells(1, 1, 3));

  assert.deepEqual(worksheet.getValues(1, 1, 3, 1), [["b c", "", ""]]);
  assert.equal(worksheet.getFormattings(1, 1, 1, 1)[0][0].italic, true);
  assert.equal(worksheet.getFormattings(1, 1, 1, 1)[0][0].bold, false);
});

test("undo of merge restores inner borders", async () => {
  const worksheet = new Worksheet();
  const history = new WorksheetHistory(worksheet);

  worksheet.edit((edit) =>
    edit.setBorder(
      { x: 1, y: 1, width: 3, height: 1 },
      new Set(["outside", "inside"]),
    ),
  );
  await Promise.resolve();
  worksheet.edit((edit) => edit.mergeCells(1, 1, 3));
  assert.equal(worksheet.getBorder(2, 1).left, false);
  await Promise.resolve();
  history.undo();

  assert.equal(worksheet.getBorder(2, 1).left, true);
  assert.equal(worksheet.getBorder(3, 1).left, true);
  assert.deepEqual(worksheet.getMerges(1, 1, 1, 1), []);
  history.redo();
  assert.equal(worksheet.getBorder(2, 1).left, false);
});

test("CSV keeps merged content in the first cell", () => {
  const worksheet = new Worksheet();

  worksheet.edit((edit) => edit.setValues([["a", "b", "c", "e"]], 1, 1));
  worksheet.edit((edit) => edit.mergeCells(1, 1, 3));

  assert.equal(worksheet.toCSV().split(/\r?\n/)[1], ",a b c,,,e");
});

test("setBorder is undoable", async () => {
  const worksheet = new Worksheet();
  const history = new WorksheetHistory(worksheet);
  const range = { x: 1, y: 1, width: 2, height: 2 };

  worksheet.edit((edit) =>
    edit.setBorder(range, new Set(["outside", "inside"])),
  );
  await Promise.resolve();
  worksheet.edit((edit) => edit.setBorder(range, new Set(["inside"]), false));
  await Promise.resolve();

  assert.equal(worksheet.getBorder(2, 1).left, false);
  history.undo();
  assert.equal(worksheet.getBorder(2, 1).left, true);
  history.undo();
  assert.deepEqual(worksheet.getBorder(1, 1), {
    top: false,
    right: false,
    bottom: false,
    left: false,
  });
  assert.equal(history.undo(), false);
  history.redo();
  assert.deepEqual(worksheet.getBorder(1, 1), {
    top: true,
    right: true,
    bottom: true,
    left: true,
  });
});

test("mergeCells merges vertically", () => {
  const worksheet = new Worksheet();

  worksheet.edit((edit) => edit.setValues([["a"], ["b"], ["c"]], 1, 1));
  worksheet.edit((edit) => edit.mergeCells(1, 1, 1, 3));

  assert.deepEqual(worksheet.getValues(1, 1, 1, 3), [["a b c"], [""], [""]]);
  assert.deepEqual(worksheet.getMerges(1, 3, 1, 1), [
    { x: 1, y: 1, width: 1, height: 3 },
  ]);
  assert.deepEqual(worksheet.getMerges(1, 4, 1, 1), []);
});

test("mergeCells merges a rectangle and removes inner borders", () => {
  const worksheet = new Worksheet();
  const range = { x: 0, y: 0, width: 2, height: 2 };

  worksheet.edit((edit) =>
    edit.setValues([
      ["a", "b"],
      ["c", "d"],
    ]),
  );
  worksheet.edit((edit) =>
    edit.setBorder(range, new Set(["outside", "inside"])),
  );
  worksheet.edit((edit) => edit.mergeCells(0, 0, 2, 2));

  assert.deepEqual(worksheet.getValues(0, 0, 2, 2), [
    ["a b c d", ""],
    ["", ""],
  ]);
  assert.deepEqual(worksheet.getMerges(1, 1, 1, 1), [
    { x: 0, y: 0, width: 2, height: 2 },
  ]);
  assert.deepEqual(worksheet.getBorder(0, 0), {
    top: true,
    right: false,
    bottom: false,
    left: true,
  });
  assert.deepEqual(worksheet.getBorder(1, 1), {
    top: false,
    right: true,
    bottom: true,
    left: false,
  });

  worksheet.edit((edit) => edit.setBorder(range, new Set(["inside"])));
  assert.equal(worksheet.getBorder(0, 1).top, false);
});

test("overlapping rectangular merges are combined", () => {
  const worksheet = new Worksheet();

  worksheet.edit((edit) => edit.mergeCells(0, 0, 2, 2));
  worksheet.edit((edit) => edit.mergeCells(1, 1, 2, 2));
  worksheet.edit((edit) => edit.mergeCells(3, 0, 1, 2));
  worksheet.edit((edit) => edit.mergeCells(2, 1, 2, 1));

  assert.deepEqual(worksheet.getMerges(), [
    { x: 0, y: 0, width: 4, height: 3 },
  ]);
});

test("getMerges returns merges intersecting a range", () => {
  const worksheet = new Worksheet();

  worksheet.edit((edit) => edit.mergeCells(1, 1, 2, 2));
  worksheet.edit((edit) => edit.mergeCells(5, 5, 1, 3));

  assert.deepEqual(worksheet.getMerges(0, 2, 10, 1), [
    { x: 1, y: 1, width: 2, height: 2 },
  ]);
  assert.deepEqual(worksheet.getMerges(5, 7, 1, 1), [
    { x: 5, y: 5, width: 1, height: 3 },
  ]);
  assert.deepEqual(worksheet.getMerges(3, 0, 2, 10), []);
});

test("rectangular merge is undoable", async () => {
  const worksheet = new Worksheet();
  const history = new WorksheetHistory(worksheet);

  worksheet.edit((edit) => edit.mergeCells(0, 0, 2, 1));
  await Promise.resolve();
  worksheet.edit((edit) => edit.mergeCells(1, 0, 1, 3));

  assert.deepEqual(worksheet.getMerges(), [
    { x: 0, y: 0, width: 2, height: 3 },
  ]);
  history.undo();
  assert.deepEqual(worksheet.getMerges(), [
    { x: 0, y: 0, width: 2, height: 1 },
  ]);
  history.redo();
  assert.deepEqual(worksheet.getMerges(), [
    { x: 0, y: 0, width: 2, height: 3 },
  ]);
});

test("merge change has the new and previous size of the merged cell", () => {
  const worksheet = new Worksheet();
  const events = [];

  worksheet.addEventListener("change", ({ changes }) =>
    events.push(
      ...changes.map(({ x, y, merge, previousMerge }) => ({
        x,
        y,
        ...(merge || previousMerge),
      })),
    ),
  );
  worksheet.edit((edit) => edit.mergeCells(1, 2, 3, 4));
  worksheet.edit((edit) => edit.unmergeCells(2, 3));

  assert.deepEqual(events, [
    { x: 1, y: 2, width: 3, height: 4 },
    { x: 1, y: 2, width: 3, height: 4 },
  ]);
});
