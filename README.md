# Codex で Power Apps キャンバスアプリを検証する環境

Microsoft 公式 Canvas Authoring MCP と公式ガイドを利用する、検証用環境です。アプリの業務機能や本番環境は変更しません。

このリポジトリのスキルは **公式資料を参考にした Codex 向け設定** です。Microsoft 公式プラグイン全体の Codex 互換性を保証するものではありません。

## 最初に読むもの

1. [セットアップ](docs/setup.md)：必要なツール、個人設定と共有設定
2. [Studio と MCP 接続](docs/canvas-mcp.md)：人が行う操作、読み取り専用の接続確認、再認証
3. [Playwright](docs/playwright.md)：サインイン、表示確認、非更新操作、テストの拡張
4. [公式資料・ライセンス・互換性](docs/sources.md)
5. [今回の確認結果](docs/verification.md)：設定と接続・実行の区別

## 作業開始

```sh
npm ci
npx playwright install chromium
npm run check
npm run test:local
npm run test:list
```

`test:local` はダミー HTML に対するブラウザー起動確認です。Power Apps に接続しません。
`test:list` はテスト列挙のみです。実アプリの成功を意味しません。

実アプリの準備後は `.local/app.json` を設定して `npm run auth` → `npm test` の順に実行します。
未設定のまま `npm test` を実行すると明示的に失敗します。接続できないテストを成功扱いにはしません。

## 保管方針

Git 管理するのは設定例、テストコード、ロックファイル、スキル、手順と匿名化した確認結果です。
実 URL・ID、認証状態、画面キャプチャ、同期した YAML、ログは `.local/` に保存し Git 管理しません。
パスワードやアクセストークンを設定例やチャットに入力しないでください。
