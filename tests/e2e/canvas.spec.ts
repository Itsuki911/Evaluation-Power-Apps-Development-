import { test, expect } from '@playwright/test';
import { loadAppConfig, appScope } from '../../scripts/app-config.mjs';
test('検証用キャンバス画面が表示される', async ({ page }) => {
  const c = loadAppConfig();
  await page.goto(c.playUrl, { waitUntil: 'domcontentloaded' });
  await expect(appScope(page, c).getByText(c.readyText, { exact: true })).toBeVisible({ timeout: 60000 });
});
test('明示した非更新操作で期待する表示になる', async ({ page }) => {
  const c = loadAppConfig();
  test.skip(!c.readOnlyAction, '実画面で非更新と確認した操作が未設定です。');
  await page.goto(c.playUrl, { waitUntil: 'domcontentloaded' });
  const app = appScope(page, c);
  await expect(app.getByText(c.readyText, { exact: true })).toBeVisible({ timeout: 60000 });
  await app.getByRole('button', { name: c.readOnlyAction.buttonName, exact: true }).click();
  await expect(app.getByText(c.readOnlyAction.expectedText, { exact: true })).toBeVisible();
});
