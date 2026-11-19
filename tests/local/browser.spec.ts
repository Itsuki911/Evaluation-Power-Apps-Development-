import { test, expect } from '@playwright/test';
test('ローカルの Chromium で表示と非更新ボタン操作ができる', async ({ page }) => {
  await page.setContent('<button onclick="document.querySelector(\'output\').textContent=\'表示確認済み\'">表示確認</button><output></output>');
  await page.getByRole('button', { name: '表示確認', exact: true }).click();
  await expect(page.locator('output')).toHaveText('表示確認済み');
});
