# 公式資料、互換性、更新

確認日：2026-11-19。リンク先は変化するため、再実行時には公開設定と実際のツールスキーマを確認します。

## Microsoft Canvas Apps

- [外部ツールによるキャンバスアプリの作成](https://learn.microsoft.com/en-us/power-apps/maker/canvas-apps/create-canvas-external-tools)
- [公式プラグイン](https://github.com/microsoft/power-platform-skills/tree/main/plugins/canvas-apps)
- 調査したコミット：`42a94220bad33ffc815c3b8fb9e7229e65ad3542`、プラグイン 3.1.1。
- [公開 MCP 起動設定](https://github.com/microsoft/power-platform-skills/blob/42a94220bad33ffc815c3b8fb9e7229e65ad3542/plugins/canvas-apps/.mcp.json)
- [接続スキル](https://github.com/microsoft/power-platform-skills/blob/42a94220bad33ffc815c3b8fb9e7229e65ad3542/plugins/canvas-apps/skills/configure-canvas-mcp/SKILL.md)

必要な作業の前に次の原文を参照します。主要な Canvas ガイドは同じ固定版を `.agents/skills/canvas-codex/references/microsoft/` に原文のまま同梱し、対応する MIT ライセンスを保存しています。以下は上流へのリンクです。

| 用途 | 固定リビジョンのガイド |
|---|---|
| YAML 構造と式の引用 | [YamlSyntax.md](https://github.com/microsoft/power-platform-skills/blob/42a94220bad33ffc815c3b8fb9e7229e65ad3542/plugins/canvas-apps/references/YamlSyntax.md) |
| コントロール契約と列挙値 | [ControlGuide.md](https://github.com/microsoft/power-platform-skills/blob/42a94220bad33ffc815c3b8fb9e7229e65ad3542/plugins/canvas-apps/references/ControlGuide.md) |
| サイズ・スクロール・レスポンシブ | [LayoutGuide.md](https://github.com/microsoft/power-platform-skills/blob/42a94220bad33ffc815c3b8fb9e7229e65ad3542/plugins/canvas-apps/references/LayoutGuide.md) |
| 状態・イベント・データ | [PowerFxGuide.md](https://github.com/microsoft/power-platform-skills/blob/42a94220bad33ffc815c3b8fb9e7229e65ad3542/plugins/canvas-apps/references/PowerFxGuide.md) |
| コンパイルで分からない実行時問題 | [QAChecks.md](https://github.com/microsoft/power-platform-skills/blob/42a94220bad33ffc815c3b8fb9e7229e65ad3542/plugins/canvas-apps/references/QAChecks.md) |

公式スキルは、接続、既存／新規アプリ編集、Studio での接続追加を支援します。依存先は Canvas Authoring MCP、Studio の共同編集、.NET 10、編集権限です。
公式プラグインの `Task`、専用エージェント、`${PLUGIN_ROOT}`、Claude/Copilot 用 hooks は今回の Codex で互換動作を確認できませんでした。プラグイン全体を無変更で導入したとは扱いません。
Codex 向けの小さなスキルで、同期の確認と公式ガイド参照を明示しました。公式の並列専門エージェントや全ワークフローを移植したものではありません。

## Playwright

- [Microsoft Learn 概要](https://learn.microsoft.com/en-us/power-platform/developer/playwright-samples/overview)
- [Microsoft 公式サンプル](https://github.com/microsoft/power-platform-playwright-samples)
- 調査コミット：`3233f3e73e30960e8ab7349f1cb06b520ca010e4`。
- [Canvas fixture](https://github.com/microsoft/power-platform-playwright-samples/blob/3233f3e73e30960e8ab7349f1cb06b520ca010e4/packages/e2e-tests/fixtures/canvas.fixtures.ts)
- [Playwright 認証](https://playwright.dev/docs/auth)、[導入](https://playwright.dev/docs/intro)

サンプルは toolkit と Northwind の E2E 実装を含みます。本リポジトリでは設計を参考にし、`@playwright/test` の小さなテストを独自作成しました。公式 toolkit と同一の網羅性はありません。

## Codex と .NET

- [Codex MCP](https://developers.openai.com/codex/mcp)
- [Codex Skills](https://developers.openai.com/codex/skills)
- [Codex CLI ソース](https://github.com/openai/codex)
- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)
- [公式 dotnet-install スクリプト](https://learn.microsoft.com/en-us/dotnet/core/tools/dotnet-install-script)

Codex CLI のローカル `mcp add --help` と Context7 の公式ソース資料で stdio 登録を照合しました。Skills は標準 frontmatter とプロジェクトの `.agents/skills/` を使います。新しい Codex セッションでの自動選択は別途確認が必要です。

## ライセンスと更新手順

両 Microsoft リポジトリの LICENSE は MIT でした。Canvas 公式原文の一部を同梱し、Codex 向けスキルは独自調整として区別しています。Microsoft の著作権・ライセンス文は [THIRD_PARTY_NOTICES.md](../THIRD_PARTY_NOTICES.md) と同梱ディレクトリに保持します。

1. 公式リポジトリを別の一時ディレクトリに取得し、上記コミットからの差分を読む。
2. `.mcp.json`、connect の入力、スキル依存、QA、ライセンスの変更を確認する。
3. 採用する変更のみ作業ブランチへ反映し、参照コミットを更新する。公式原文を独自変更した場合はその旨を記録する。
4. MCP probe、型検査、ローカルブラウザー確認を実行する。
5. 検証用 Studio に読み取り接続し、アプリの実テストを別途実施する。

`dnx` のパッケージは公式公開設定どおりバージョン未固定です。今回起動した版は確認結果に記録しています。SDK・NuGet パッケージ・npm 依存を同時に更新せず、変更ごとに検証してください。

2026-11-19 に `git ls-remote` で上流 main が上記コミットと一致することを確認しました。同梱ファイルは自動更新されません。更新時は上流の同じパスを比較し、採用した版の原文・ライセンス・コミット ID を一緒に更新します。公式クイックインストーラーは Copilot／Claude 向けであり Codex への全面対応を確認できないため実行していません。
