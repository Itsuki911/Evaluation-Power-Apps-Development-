# セットアップ

## 前提

- Git、Node.js 22 以上と npm、Codex CLI またはローカル MCP を起動できる Codex アプリ。
- .NET **SDK** 10 以上と `dnx`。ランタイムだけでは不足します。
- 組織アカウント、利用可能な Power Apps ライセンス、対象環境・検証用アプリの編集権限。組織の DLP・接続ポリシーが適用されます。
- 保存済みの検証用キャンバスアプリと Studio の共同編集セッション。

初回は `git status --short --branch` と既存の `AGENTS.md` を確認し、作業ブランチを作ります。
GitHub 認証が必要なら Git Credential Manager または GitHub CLI の対話型 `gh auth login` を人が実行します。トークンを remote URL に埋め込まないでください。

## .NET

通常の開発機では [Microsoft の .NET 10 ダウンロード](https://dotnet.microsoft.com/download/dotnet/10.0) から OS・CPU に合う SDK を導入します。
今回の macOS arm64 では、システム全体を変更しない評価用の方法として、公式スクリプトでユーザー専用の SDK を導入しました。

```sh
curl -fsSL https://dot.net/v1/dotnet-install.sh -o /tmp/canvas-dotnet-install.sh
bash /tmp/canvas-dotnet-install.sh --channel 10.0 --quality GA --architecture arm64 \
  --install-dir "$HOME/.local/share/canvas-dotnet" --no-path
export PATH="$HOME/.local/share/canvas-dotnet:$PATH"
dotnet --list-sdks
dnx --help
```

`export` はその端末だけに作用します。Codex アプリには `dnx` の絶対パスを登録するため、シェル設定の変更は不要です。
公式スクリプトは主に CI 向けです。継続利用では SDK 更新を管理するか、公式インストーラーへ移行してください。
SDK 初回起動には .NET の初回セットアップ処理があります。HTTPS 証明書の信頼設定は本環境では不要です。

## Codex の MCP 登録

`~/.codex/config.toml` を丸ごと置き換えないでください。既存の `canvas-authoring` があれば、先に設定を確認します。

```sh
codex mcp get canvas-authoring
# 未登録の場合のみ実行。dnx の場所は実機に合わせる。
codex mcp add canvas-authoring -- "$HOME/.local/share/canvas-dotnet/dnx" \
  Microsoft.PowerApps.CanvasAuthoring.McpServer --yes
codex mcp get canvas-authoring
```

登録コマンドはユーザー設定を更新します。公開設定例は `config/codex-mcp.example.toml` です。
初回確認用の `enabled_tools` とタイムアウト設定を既存の同名セクションへ追加できます。二重の TOML セクションを作らないでください。
プロジェクト単位なら `.codex/config.toml` に設定できますが、Codex に信頼されたプロジェクトでのみ読み込まれます。実機では未信頼のため最初は検出されず、ユーザー設定への登録に切り替えました。

**登録後は Codex を再起動するか新しいセッションを開始**し、`canvas-authoring` のツールが公開されたことを確認します。現在の会話に後から自動追加されるとは限りません。
このリポジトリを Codex の作業場所として開くと `.agents/skills/canvas-codex/SKILL.md` を利用できます。見つからなければファイルを明示して読むよう依頼してください。

## サーバー起動だけを確認

```sh
CANVAS_DNX="$HOME/.local/share/canvas-dotnet/dnx" npm run mcp:probe
```

このスクリプトは MCP の `initialize` と `tools/list` のみを実行します。アプリへの接続や認証はしません。
NuGet へ初回アクセスできることが必要です。`--yes` は公式公開設定に従います。固定バージョンを使いたい場合は `dnx Package@version` を使い、更新時に再検証してください。

## Playwright

```sh
npm ci
npx playwright install chromium
npm run check
npm run test:local
```

Playwright MCP は不要です。既存のブラウザー操作機能と、独立した Playwright Test を使います。
`package-lock.json` で依存バージョンを固定します。更新時には型検査・ローカル確認・検証用アプリ確認を分けて実施してください。
