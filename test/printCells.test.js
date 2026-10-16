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

test('cellsToSVG draws borders centered on the grid lines', () => {
  const { svg, width, height } = cellsToSVG([['a']], {
    getBorder: (column, row) => (column === 0 && row === 0 ? { top: true, right: true, bottom: true, left: true } : undefined)
  });

  assert.equal(width, 103);
  assert.equal(height, 27);
  assert.equal((svg.match(/<line [^>]*stroke-width="3"/g) || []).length, 4);
  assert.ok(svg.includes('<line x1="1.5" y1="1.5" x2="1.5" y2="25.5"'));
  assert.ok(svg.includes('<rect x="1.5" y="1.5"'));
});

test('cellsToSVG draws borders set on neighbors outside of the cells', () => {
  const { svg } = cellsToSVG([['a']], { getBorder: (column, row) => (column === -1 && row === 0 ? { right: true } : undefined) });

  assert.ok(svg.includes('<line x1="1.5" y1="1.5" x2="1.5" y2="25.5"'));
});
