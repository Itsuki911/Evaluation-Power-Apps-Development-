# Studio と Canvas Authoring MCP

## 人が Studio で行う操作

1. `make.powerapps.com` に組織アカウントでサインインし、検証に使ってよい環境を選びます。
2. 空の検証用キャンバスアプリを作成して保存します。既存の検証用アプリでも構いません。
3. Studio の **Settings → Updates → Coauthoring** を有効にします。設定が表示されない場合は利用条件・組織ポリシーを管理者に確認します。
4. Studio のタブを開いたままにし、完全な Studio URL を Codex に指定します。公開リポジトリには保存しません。
5. MCP から認証要求が出たら、Studio と同じアカウントでサインイン・MFA を完了します。新たな管理者承認が必要なら管理者へ依頼します。

## 接続確認の順序

新しい Codex セッションでこのように依頼します。

> canvas-codex の手順を読み、指定した検証用 Studio URL に Canvas Authoring MCP で接続してください。list_controls と list_data_sources で読み取り確認だけ行い、compile_canvas、アプリ変更、公開、実データの取得・更新はしないでください。

実際の `connect` の入力スキーマを確認し、次の値だけを渡します。

- `environment_id`：Studio URL の `/e/` の次のパス要素。
- `app_id`：`app-id` クエリを URL デコードした最後のパス要素。
- `environment_category`：ホストが `make.powerapps.com` または `make.preview.powerapps.com` なら `prod`。これはクラウド区分で、検証／本番のアプリ用途とは別です。他クラウドは公式 configure-canvas-mcp の対応表を確認します。

`login_hint`、`tenant_id`、認証方式は通常省略します。URL から推測しません。

接続成功後、`list_controls` → 必要な `describe_control` → `list_data_sources` → 接続済みソースの `get_data_source_schema` の順に確認します。データソースがゼロならそのまま記録し、取得失敗と区別します。
構成確認には `sync_canvas` を使えますが、空の `.local/canvas/<app>/` を出力先に指定してください。ファイルの上書きがあるため既存の作業ディレクトリに無確認で同期しません。

**`compile_canvas` は YAML 検証と live アプリへの適用を行います。接続テストでは呼び出しません。**
将来の編集では変更範囲を確認し、現在のアプリを同期、公式ガイドとメタデータに基づき YAML/Power Fx を編集、差分を確認してから適用します。適用後は Studio のプレビューと Playwright を別途検証します。

## データソース

SharePoint の検証用リスト、Dataverse の検証用テーブル、OneDrive/SharePoint 上の検証用 Excel テーブルなどを想定しています。
新規接続は人が Studio の **Data / Add data** から追加し、必要なコネクターにサインインします。列を作る場合も検証用データだけに限定します。
その後 MCP で一覧と列名・型を再取得します。接続追加と既存データ行の読み取りは別です。
今回 Dataverse MCP は追加しません。スキーマ取得には Canvas Authoring MCP を使い、別の MCP が必要になったら用途・対象データ・必要権限を先に説明します。

## 再接続とトラブルシューティング

| 症状 | 確認・対処 |
|---|---|
| `dnx` がない | SDK 10 以上と絶対パスを確認。ランタイムのみでは不可 |
| NuGet の接続失敗 | `api.nuget.org` への DNS、プロキシ、組織の許可を確認 |
| Codex にツールが出ない | `codex mcp get canvas-authoring`、プロジェクト信頼、Codex 再起動。別ホスト用のツール接頭辞を探さない |
| 起動タイムアウト | 初回パッケージ取得を probe で済ませ、`startup_timeout_sec=120` を設定 |
| 接続 404 | Studio と同じアカウント・正しい URL・開いたタブを確認。`force_account_select=true` で再試行、失敗時は `auth_flow=browser` も指定。2回の回復失敗後は停止して原因を報告 |
| 401/403 | MFA、環境とアプリ権限、組織ポリシーを確認。必要ならアカウント選択を明示。権限を自動拡張しない |
| B2B ゲスト | 人が確認したリソース側 tenant_id を使用。推測しない |
| セッション切れ | Studio を開き直し、共同編集を確認して `connect` を再実行 |
| データソースが見えない | Studio 側の接続・テーブル追加を確認して一覧／スキーマを再取得 |
| 認証方式が動かない | browser/broker/devicecode は実スキーマと公式手順に従う。devicecode は MCP elicitation に対応するクライアントが必要 |

ログやサポートバンドルは認証・機密情報を含み得ます。今回 `create_support_bundle` は使用しません。
