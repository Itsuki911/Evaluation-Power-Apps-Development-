---
name: canvas-codex
description: Power Apps の検証用キャンバスアプリで Canvas Authoring MCP の接続確認、pa.yaml と Power Fx の編集準備、ブラウザー QA を支援する。
---

# Canvas Apps を Codex で扱う

これは Microsoft 公式プラグインそのものではなく、公式資料を参考にした Codex 向け設定。
参照元と固定リビジョンは `docs/sources.md`、接続手順は `docs/canvas-mcp.md` を読む。

- 最初に対象が検証用であること、Studio の保存・共同編集・開いたタブを確認する。
- 実際に利用可能な Canvas Authoring MCP のツールスキーマを調べる。Claude の `Task`、`AskUserQuestion`、`${PLUGIN_ROOT}`、フック名を Codex で使えると仮定しない。
- `connect` の成功後、`list_controls`、`list_data_sources`、必要な `describe_control` と `get_data_source_schema` でメタデータだけ確認する。レコード取得や更新は今回の範囲外。
- ID は Studio URL から取得する。`environment_category=prod` は公開クラウドの区分であり、業務用本番アプリを指定する意味ではない。テスト用途という理由で `test` を指定しない。
- アプリ構成を読む場合、`sync_canvas` の出力は `.local/canvas/<app>/` の空の専用ディレクトリにする。リポジトリ直下を渡さない。同期結果は機密を確認するまで Git 管理しない。
- 編集依頼がある場合のみ、同期済み YAML と以下の公式原文を対象に応じて読む。原文中の他ホスト向けツール名・承認フローはこの環境に自動適用せず、実際に利用可能な機能とユーザーの指示を優先する。コントロールの型、バージョン、Variant、列の型、列挙値は MCP の応答を使う。
- YAML のプロパティ式は `=` で始め、コロンを含む式は YAML として引用する。既存のコントロール名・データバインド・対象外のプロパティを保持する。
- 変更前に再同期してよいかローカル差分を確認する。未適用の編集を上書きしない。
- `compile_canvas` は検証と適用を行う。接続確認のためには呼ばない。将来、対象と変更範囲が承認された検証用アプリでのみ使用する。
- コンパイル成功は表示・操作の成功を意味しない。公式 QA ガイドの適用項目と、実ブラウザーでのスクロール、狭い幅、ラベル、操作結果を別に検証する。
- 読み取り専用テストでも OnStart/OnVisible/Timer に副作用がないか確認する。保存・削除・本番データのテストは行わない。
- 最終報告では設定、起動、接続、メタデータ取得、コンパイル、ブラウザー操作、自動テストの証拠を分ける。

## 公式ガイドの選択

原文は Microsoft の固定コミット `42a94220bad33ffc815c3b8fb9e7229e65ad3542` の MIT ライセンス資料。下記を必要に応じて開く。全文を毎回読む必要はない。

- YAML 作成・修復：[YamlSyntax](references/microsoft/YamlSyntax.md)、[ControlGuide](references/microsoft/ControlGuide.md)。`describe_control` の作成用キーワード・Variant と照合する。
- 画面配置：[LayoutGuide](references/microsoft/LayoutGuide.md)、GridLayout が対象なら [GridLayoutGuide](references/microsoft/GridLayoutGuide.md)、見た目の設計なら [DesignGuide](references/microsoft/DesignGuide.md)。
- 状態・イベント・式：[PowerFxGuide](references/microsoft/PowerFxGuide.md)。データ操作は要求された検証用データの範囲だけにする。
- 生成後：[QAChecks](references/microsoft/QAChecks.md) の目次から変更と依存箇所に該当するチェックを選び、コントロール契約、式、重複キー、アクセシビリティ、狭い幅、スクロールと操作結果を確認。未実行項目を明示する。

生成前に必要な画面・操作・状態・データを整理し、不明な接続や制約は推測せず確認する。既存アプリは同期結果に基づき対象外を保持する。生成結果を適用する場合は、利用可能な MCP 操作と作用を先に確認する。
