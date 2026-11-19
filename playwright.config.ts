import { defineConfig } from '@playwright/test';
import { existsSync } from 'node:fs';
import { authFile, loadAppConfig } from './scripts/app-config.mjs';
const listing = process.argv.includes('--list');
if (!listing) {
  loadAppConfig();
  if (!existsSync(authFile)) throw new Error('npm run auth で検証用アカウントにサインインしてください。');
}
export default defineConfig({
  testDir: './tests/e2e',
  workers: 1,
  retries: 0,
  timeout: 90000,
  reporter: [['list']],
  outputDir: '.local/test-results',
  use: {
    browserName: 'chromium',
    storageState: listing ? undefined : authFile,
    screenshot: 'off', video: 'off', trace: 'off',
    actionTimeout: 15000,
  },
  projects: [
    { name: 'desktop', use: { viewport: { width: 1440, height: 900 } } },
    { name: 'narrow', use: { viewport: { width: 390, height: 844 } } },
  ],
});
