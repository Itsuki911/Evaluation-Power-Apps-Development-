import { chromium } from '@playwright/test';
import { mkdir, chmod } from 'node:fs/promises';
import { createInterface } from 'node:readline/promises';
import { loadAppConfig, appScope, authFile, localDir } from './app-config.mjs';
const c = loadAppConfig();
await mkdir(localDir, { recursive: true, mode: 0o700 });
await chmod(localDir, 0o700);
process.umask(0o077);
const browser = await chromium.launch({ headless: false });
const input = createInterface({ input: process.stdin, output: process.stdout });
try {
  const context = await browser.newContext();
  const page = await context.newPage();
  await page.goto(c.playUrl, { waitUntil: 'domcontentloaded' });
  await input.question('ブラウザーでサインイン・MFA を完了し、検証画面が表示されたら Enter（認証情報は入力しない）: ');
  await appScope(page, c).getByText(c.readyText, { exact: true }).waitFor({ state: 'visible', timeout: 30000 });
  await context.storageState({ path: authFile, indexedDB: true });
  await chmod(authFile, 0o600);
  console.log('認証状態を Git 除外済みの .local/auth.json に保存しました。');
} finally {
  input.close();
  await browser.close();
}
