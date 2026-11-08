import assert from 'node:assert';
import test from 'node:test';

import { formatDisplayValue } from '../src/worksheet/numberFormats.ts';

test('formatDisplayValue keeps values as-is without a format or for non-numbers', () => {
  assert.deepStrictEqual(formatDisplayValue(1.5, undefined), { text: '1.5', negative: false });
  assert.deepStrictEqual(formatDisplayValue('abc', 'currency'), { text: 'abc', negative: false });
  assert.deepStrictEqual(formatDisplayValue(true, 'percent'), { text: 'TRUE', negative: false });
});

test('formatDisplayValue formats currency with 2 decimal places, sign is separated', () => {
  const expected = (1234.5).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  assert.deepStrictEqual(formatDisplayValue(1234.5, 'currency'), { text: expected, negative: false });
  assert.deepStrictEqual(formatDisplayValue(-1234.5, 'currency'), { text: expected, negative: true });
  assert.strictEqual(formatDisplayValue(-0.001, 'currency').negative, false);
});

test('formatDisplayValue formats percentage with 1 decimal place', () => {
  assert.strictEqual(formatDisplayValue(0.1234, 'percent').text, (12.34).toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 }));
});

test('formatDisplayValue formats date and date/time, rounding down milliseconds', () => {
  assert.strictEqual(formatDisplayValue(0, 'date').text, new Date(0).toLocaleDateString());
  assert.strictEqual(formatDisplayValue(1999, 'datetime').text, new Date(1000).toLocaleString());
  assert.strictEqual(formatDisplayValue(1e20, 'date').text, formatDisplayValue(1e20, undefined).text);
});

test('formatDisplayValue formats duration as HH:MM:SS or MM:SS, rounding down milliseconds', () => {
  assert.strictEqual(formatDisplayValue(65_999, 'duration').text, '01:05');
  assert.strictEqual(formatDisplayValue(3_723_000, 'duration').text, '01:02:03');
  assert.strictEqual(formatDisplayValue(-61_000, 'duration').text, '-01:01');
});
