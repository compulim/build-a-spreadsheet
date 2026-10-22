import assert from "node:assert/strict";
import { test } from "node:test";
import { calculate } from "../src/calculator.js";

test("calculates valid expressions", () => {
  assert.equal(calculate("123+456"), 579);
  assert.equal(calculate("1 + 2 * 3"), 7);
  assert.equal(calculate("((1+2)*(3-1))/4"), 1.5);
  assert.equal(calculate("-(2+3)*-2"), 10);
  assert.equal(calculate("10-2-3"), 5);
  assert.equal(calculate("8/2/2"), 2);
  assert.equal(calculate(".5*1e2"), 50);
});

test("returns undefined for invalid expressions", () => {
  for (const input of [
    "",
    "1+",
    "(1+2",
    "1+2)",
    "abc",
    "1/0",
    "1 2",
    "SUM(1)",
  ]) {
    assert.equal(calculate(input), undefined, input);
  }
});
