import { defineConfig } from '@playwright/test';
export default defineConfig({
  testDir: './tests/local', workers: 1, retries: 0,
  outputDir: '.local/local-test-results', reporter: 'list',
  use: { browserName: 'chromium', screenshot: 'off', video: 'off', trace: 'off' },
});
