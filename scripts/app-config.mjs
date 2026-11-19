import fs from 'node:fs';
import path from 'node:path';
export const localDir = path.resolve('.local');
export const authFile = path.join(localDir, 'auth.json');
export function loadAppConfig() {
  const file = path.join(localDir, 'app.json');
  if (!fs.existsSync(file)) throw new Error('config/app.example.json を .local/app.json にコピーして検証用設定を入力してください。');
  const c = JSON.parse(fs.readFileSync(file, 'utf8'));
  if (c.testOnly !== true || c.startupReviewed !== true) throw new Error('検証用アプリであり、OnStart/OnVisible が実データを更新しないことの確認が必要です。');
  const url = new URL(c.playUrl);
  if (url.protocol !== 'https:' || url.username || url.password || !/^apps\.powerapps\.(com|us|cn)$/.test(url.hostname) || !url.pathname.startsWith('/play/')) throw new Error('Studio URL ではなく、検証用アプリの HTTPS 再生 URL を指定してください。');
  if (!c.readyText?.trim()) throw new Error('readyText に検証画面内の固有の表示文字列を指定してください。');
  if (!Array.isArray(c.frameSelectors) || c.frameSelectors.some(s => typeof s !== 'string' || !s.trim())) throw new Error('frameSelectors は実画面で確認した iframe セレクターの配列にしてください。');
  if (c.readOnlyAction !== null && (!c.readOnlyAction || c.readOnlyAction.confirmedNonMutating !== true || !c.readOnlyAction.buttonName?.trim() || !c.readOnlyAction.expectedText?.trim())) throw new Error('操作テストには非更新操作の確認、buttonName、expectedText が必要です。');
  return c;
}
export function appScope(page, config) {
  return config.frameSelectors.reduce((scope, selector) => scope.frameLocator(selector), page);
}
