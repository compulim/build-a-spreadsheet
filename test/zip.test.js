import assert from 'node:assert/strict';
import { test } from 'node:test';

import { crc32, createZip, readZip } from '../src/zip.ts';

const encoder = new TextEncoder();

test('crc32', () => {
  assert.equal(crc32(encoder.encode('123456789')), 0xcbf43926);
});

for (const compress of [true, false]) {
  test(`round trip (compress: ${compress})`, async () => {
    const entries = [
      { name: 'Sheet1.csv', data: encoder.encode('a,b\r\n'.repeat(1000)) },
      { name: 'Empty.csv', data: new Uint8Array() },
      { name: 'Unicode 日本語.csv', data: encoder.encode('こんにちは') }
    ];

    const zip = await createZip(entries, { compress });
    const actual = await readZip(zip);

    assert.deepEqual(
      actual.map(({ name, data }) => ({ name, data: [...data] })),
      entries.map(({ name, data }) => ({ name, data: [...data] }))
    );

    if (compress) {
      assert.ok(zip.length < 5000);
    }
  });
}

test('invalid ZIP', async () => {
  await assert.rejects(readZip(encoder.encode('not a zip file at all, really not')), /Invalid ZIP/);
});
