import assert from 'node:assert/strict';
import { test } from 'node:test';
import { cellsToSVG } from '../src/printCells.ts';


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

test('cellsToSVG renders formatting', () => {
  const { svg } = cellsToSVG([['a', 1, true]], {
    getFormatting: column =>
      column === 0
        ? { bold: true, italic: true, underline: true, strikethrough: true, align: 'center', highlight: '#c6efce/' }
        : column === 1
          ? { align: 'left', highlight: '#ffdddd/#f70000' }
          : undefined
  });

  assert.ok(svg.includes(`<feFlood flood-color="#c6efce"/>`));
  assert.ok(svg.includes('filter="url(#m0)"'));
  assert.equal((svg.match(/<rect [^/]*fill="#fff"/g) || []).length, 3);
  assert.ok(svg.includes('text-anchor="middle" font-weight="bold" font-style="italic" text-decoration="underline line-through"'));
  assert.ok(svg.includes(`x="104.5" y="12.5" dominant-baseline="central" fill="#f70000">1</text>`));
  assert.ok(svg.includes('text-anchor="middle">TRUE</text>'));
});

test('cellsToSVG draws merged cells as a single cell without inner borders', () => {
  const { svg } = cellsToSVG([['merged', '', 'c']], {
    getBorder: () => ({ top: true, right: true, bottom: true, left: true }),
    getColumnSpan: column => (column === 0 ? 2 : column === 1 ? 0 : 1)
  });

  assert.equal((svg.match(/<rect [^/]*fill/g) || []).length, 2);
  assert.ok(svg.includes('width="200" height="24" fill="#fff"'));
  assert.ok(!svg.includes('<line x1="101.5" y1="1.5" x2="101.5"'));
  assert.ok(svg.includes('<line x1="201.5" y1="1.5" x2="201.5"'));
});

test('cellsToSVG draws vertically merged cells as a single cell without inner borders', () => {
  const { svg } = cellsToSVG([['merged'], ['']], {
    getBorder: () => ({ top: true, right: true, bottom: true, left: true }),
    getRowSpan: (_, row) => (row === 0 ? 2 : 0)
  });

  assert.equal((svg.match(/<rect [^/]*fill/g) || []).length, 1);
  assert.ok(svg.includes('width="100" height="48" fill="#fff"'));
  assert.ok(!svg.includes('<line x1="1.5" y1="25.5" x2="101.5" y2="25.5"'));
  assert.ok(svg.includes('<line x1="1.5" y1="49.5" x2="101.5" y2="49.5"'));
});
