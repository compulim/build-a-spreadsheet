/// <reference types="node" />

import { BrowserRun, FileSystemSnapshotStore } from '@onting/harness';
import { isCI } from 'ci-info';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import createStubImplementation from '@onting/stub/host.js';

const __FILENAME = fileURLToPath(import.meta.url);
const TEST_HTML_FILENAME = 'html/merge-cells-navigate-enter.html';
const runHeadless = isCI || 'CODESPACES' in process.env;

test(TEST_HTML_FILENAME, async () => {
  const stubImplementation = await createStubImplementation(
    () => new FileSystemSnapshotStore({ testFilePath: __FILENAME })
  );

  const run = new BrowserRun(new URL(TEST_HTML_FILENAME, runHeadless ? 'http://web/' : 'http://localhost:3000/').href, {
    stubImplementation,
    testFilePath: __FILENAME,
    ...(runHeadless ? { webDriverURL: 'http://localhost:4444/wd/hub' } : {})
  });

  await run.promise;
});
