/// <reference types="node" />

import { BrowserRun, FileSystemSnapshotStore } from '@onting/harness';
import { isCI } from 'ci-info';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import createStubImplementation from '@onting/stub/host.js';
import { dirname, relative } from 'node:path';

export default function runTest(importMetaURL: string): void {
  const testFilePath = fileURLToPath(importMetaURL);
  const htmlFilename = relative(dirname(fileURLToPath(import.meta.url)), testFilePath).replace('.test.ts', '.html');
  const runHeadless = isCI || 'CODESPACES' in process.env;

  test(htmlFilename, async () => {
    const stubImplementation = await createStubImplementation(() => new FileSystemSnapshotStore({ testFilePath }));

    const run = new BrowserRun(
      new URL(htmlFilename, runHeadless ? 'http://web/html/' : 'http://localhost:3000/html/').href,
      {
        stubImplementation,
        testFilePath,
        ...(runHeadless ? { webDriverURL: 'http://localhost:4444/wd/hub' } : {})
      }
    );

    await run.promise;
  });
}
