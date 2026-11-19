# 確認結果

確認日：2026-11-19。実際に観測した結果と、未確認事項を区別しています。

## 環境

| 項目 | 結果 |
|---|---|
| OS / CPU | macOS 27.0 / arm64 |
| Codex | ローカルアプリの会話と CLI 0.170.0。CLI の MCP 登録を実確認 |
| Git | 2.53.0。対象を取得。取得時は空でブランチ・コミット・既存ファイルなし |
| Node / npm | v24.2.0 / 11.3.0 |
| .NET SDK | 10.0.404 を `~/.local/share/canvas-dotnet` に導入。`dotnet --list-sdks` と `dnx --help` 成功 |
| Playwright Test | 1.65.1、Chromium v1255 導入済み |
| TypeScript | 7.0.2 |
| 既存ブラウザー機能 | Chrome タブ一覧取得成功。タブバインドは timeoutMs エラー。ネイティブ Chrome の画面取得成功 |
| ブラウザー用 MCP | 既存 Playwright MCP あり。重複追加なし |
| CodeGraph | server が未初期化と応答。初期化の確認待ち。索引は未作成 |

## 成功を確認した範囲

| 確認 | 結果と意味 |
|---|---|
| MCP 設定 | ユーザーの Codex 設定に canvas-authoring を追加。`codex mcp get` で command、タイムアウト、ツール許可リストを確認 |
| 設定保持 | 許可リスト追加前後の TOML を比較し、他の設定が不変であることを確認。保護されたバックアップはリポジトリ外 |
| MCP 起動 | 公式 NuGet パッケージを起動。initialize 成功、protocolVersion 2024-11-05 |
| MCP ツール取得 | tools/list 成功、16 ツール。アプリ接続を伴わない |
| MCP バージョン | `CanvasAuthoringMcpServer`、`1.2.3+87235cb86f` |
| Studio の表示 | 既存 Chrome の編集画面で空の Screen1 とコンテナーを視認。再生モードや業務操作の検証ではない |
| npm 導入 | 依存導入成功、導入時監査で 0 vulnerabilities |
| 型検査 | `npm run check` 成功 |
| JS/Python 構文 | node --check と Python AST パース成功 |
| スキル構造 | skill-creator の quick_validate.py 成功（PyYAML のある既存 `python` を使用） |
| テスト列挙 | `npm run test:list` で表示・非更新操作 × 2 画面幅の 4 テストを検出 |
| ローカル実行 | `npm run test:local` 1 passed。ダミー HTML の表示とボタン操作のみ |
| 未設定時の停止 | `npm test` が app.json 未設定で明示的に失敗。実アプリテストは実行されていない |
| Git 除外 | 認証状態、実設定、同期 YAML、画像、テスト結果、.env、個人 MCP 設定が除外されることを確認 |

通常サンドボックスでは DNS/ネットワーク接続と Chromium の MachPort 起動が拒否されました。該当操作のみ承認付きで再実行し、ダウンロードとローカルテストが成功しました。ブラウザーの失敗を Power Apps の不具合とは扱いません。

## 未確認・人の操作待ち

- 検証対象として既存 Test App を使うことの確認、共同編集の有効化確認。
- Codex を再起動／新規セッションで開き、追加した MCP とスキルが会話から利用可能かの確認。
- `connect` と Microsoft 認証・MFA。MCP は **Power Apps 未接続**。
- アプリのコントロール一覧、接続済みデータソース、列スキーマの実取得。
- `sync_canvas` による実アプリ構成の取得。
- 実アプリの readyText、iframe 構造、非更新操作の確定。
- Playwright の対話認証・認証再利用、実アプリの表示・操作・複数幅でのテスト。
- YAML 生成、Power Fx 編集、`compile_canvas`。今回は範囲外なので未実行。
- Microsoft 公式プラグイン全体の Codex での hooks／専門エージェント互換性。

本番アプリ・データ・共有設定の変更、アプリ保存・公開、データ接続追加は行っていません。

## GitHub と再開

作業ブランチ：`setup/codex-canvas-environment`。
取得時点でリポジトリは完全に空で、`main` のコミットがありませんでした。通常の `main` 宛 PR は比較先を用意するまで作成できません。`main` への直接変更は行いません。
GitHub 公開の最終状態は完了報告を参照してください。ローカルのみの場合は GitHub 認証後に次で再開できます。

```sh
git push -u origin setup/codex-canvas-environment
```

空の GitHub リポジトリでは最初に push したブランチが既定ブランチになる場合があります。既定ブランチ・PR の比較先は所有者の方針に合わせて整理してください。
