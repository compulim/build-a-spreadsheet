import assert from 'node:assert/strict';
import { describe, test } from 'node:test';

import { formatValue, inferValue, parseCSV, parseDelimited, parseTSV, stringifyCSV, stringifyTSV } from '../src/csv.js';

describe('inferValue', () => {
  test('numbers', () => {
    assert.equal(inferValue('123'), 123);
    assert.equal(inferValue('-1.5'), -1.5);
    assert.equal(inferValue('.5'), 0.5);
    assert.equal(inferValue('1e3'), 1000);
  });

  test('booleans', () => {
    assert.equal(inferValue('TRUE'), true);
    assert.equal(inferValue('false'), false);
  });

  test('strings', () => {
    assert.equal(inferValue(''), '');
    assert.equal(inferValue('abc'), 'abc');
    assert.equal(inferValue('1,234'), '1,234');
    assert.equal(inferValue('$5'), '$5');
    assert.equal(inferValue('Infinity'), 'Infinity');
  });
});

test('formatValue', () => {
  assert.equal(formatValue(true), 'TRUE');
  assert.equal(formatValue(false), 'FALSE');
  assert.equal(formatValue(1.5), '1.5');
  assert.equal(formatValue('a'), 'a');
});

describe('parseDelimited', () => {
  test('empty', () => assert.deepEqual(parseDelimited(''), []));

  test('simple with trailing line break', () =>
    assert.deepEqual(parseDelimited('a,b\r\nc,d\r\n'), [
      ['a', 'b'],
      ['c', 'd']
    ]));

  test('LF and CR line breaks', () =>
    assert.deepEqual(parseDelimited('a\nb\rc'), [['a'], ['b'], ['c']]));

  test('quoted fields', () =>
    assert.deepEqual(parseDelimited('"a,b","c""d","e\r\nf"\r\n,'), [
      ['a,b', 'c"d', 'e\r\nf'],
      ['', '']
    ]));

  test('byte order mark', () => assert.deepEqual(parseDelimited('\ufeffa,b'), [['a', 'b']]));
});

test('CSV round trip', () => {
  const rows = [
    ['Name', 'Age', 'Active'],
    ['Alice, Jr.', 30, true],
    ['Bob "The Builder"', 1.5, false],
    ['Multi\nline', '', 'x']
  ];

  const csv = stringifyCSV(rows);

  assert.equal(csv.endsWith('\r\n'), true);
  assert.deepEqual(parseCSV(csv), rows);
});

test('TSV from Excel', () => {
  // Excel quotes cells containing line breaks and always ends with a line break.
  assert.deepEqual(parseTSV('1\tabc\tTRUE\r\n"line1\nline2"\t\t2.5\r\n'), [
    [1, 'abc', true],
    ['line1\nline2', '', 2.5]
  ]);

  assert.equal(
    stringifyTSV([
      [1, 'a\tb'],
      [true, 'x']
    ]),
    '1\t"a\tb"\r\nTRUE\tx'
  );
});
