import assert from 'node:assert/strict';
import { test } from 'node:test';
import { cellsToSVG } from '../src/printCells.js';

test('cellsToSVG sizes the SVG to the cells and escapes text', () => {
  const { svg, width, height } = cellsToSVG([
    ['<a&b>', 1],
    [true, undefined]
  ]);

  assert.equal(width, 201);
  assert.equal(height, 49);
  assert.match(svg, /^<svg xmlns="http:\/\/www.w3.org\/2000\/svg"/);
  assert.ok(svg.includes('&#60;a&#38;b&#62;'));
  assert.ok(svg.includes('text-anchor="end"'));
  assert.ok(svg.includes('>TRUE</text>'));
  assert.equal((svg.match(/<rect [^/]*fill/g) || []).length, 4);
});
