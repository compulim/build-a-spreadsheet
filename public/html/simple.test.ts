/// <reference types="node" />

import { BrowserRun } from '@onting/harness';
import isCI from 'is-ci';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';

const TEST_HTML_FILENAME = 'html/simple.html';

test(TEST_HTML_FILENAME, async () => {
  const run = new BrowserRun(new URL(TEST_HTML_FILENAME, isCI ? 'http://web/' : 'http://localhost:3000/').href, {
    testFilePath: fileURLToPath(import.meta.url),
    ...(isCI ? { webDriverURL: 'http://localhost:4444/wd/hub' } : {})
  });

  await run.promise;
});
