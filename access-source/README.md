# Access 勤怠管理アプリ 疑似エクスポート構成 (`access-source`)

## 1. 概要

本ディレクトリは、Microsoft Access で構築されたシンプルな従業員勤怠管理アプリ（`AttendanceManagement`）を想定し、Access のテキストエクスポート機能（`Application.SaveAsText` / VBE モジュールエクスポート / クエリ SQL 抽出）で得られる構成ファイル群を**疑似的に生成**したものです。

Microsoft Access の実行環境および `.accdb` ファイルは使用していません。

### 想定する Access アプリの主要機能

- 従業員一覧の閲覧および部署による絞り込み（`frmEmployees`）
- 在籍中（`Active`）従業員の選択と現在の勤怠状態表示（`frmMain`）
- 出勤打刻（重複打刻防止・未退勤チェック付き）（`frmMain`, `modAttendance`, `modValidation`）
- 退勤打刻（未退勤レコードの存在確認と更新）（`frmMain`, `modAttendance`, `modValidation`）
- 未退勤従業員の一覧表示（`qryOpenAttendance`, `frmMain`）
- 勤怠記録の一覧表示、従業員別履歴表示、勤務日・状態によるフィルタリング、件数集計（`frmAttendance`, `qryAttendanceHistory`）
- 従業員別の勤怠履歴レポート出力（`rptAttendanceHistory`）
- 起動時スタートアップマクロによるメイン画面自動表示（`AutoExec`）

---

## 2. ディレクトリ構成とファイル分類

本ディレクトリ内のファイルは、**Access 実エクスポートを模擬したファイル**、**実データファイル**、**独自設計の補助ファイル**、**評価用テスト定義**の4種類に明確に区別しています（詳細は `manifest.json` を参照してください）。

```text
access-source/
├── forms/
│   ├── frmMain.txt                 # [simulated_access_export] SaveAsText (acForm) 模擬
│   ├── frmEmployees.txt            # [simulated_access_export] SaveAsText (acForm) 模擬
│   └── frmAttendance.txt           # [simulated_access_export] SaveAsText (acForm) 模擬
├── reports/
│   └── rptAttendanceHistory.txt    # [simulated_access_export] SaveAsText (acReport) 模擬
├── queries/
│   ├── qryEmployees.sql            # [simulated_access_export] Access SQL クエリ定義
│   ├── qryAttendanceHistory.sql    # [simulated_access_export] Access SQL クエリ定義
│   └── qryOpenAttendance.sql       # [simulated_access_export] Access SQL クエリ定義
├── modules/
│   ├── modAttendance.bas           # [simulated_access_export] VBA 標準モジュール
│   └── modValidation.bas           # [simulated_access_export] VBA 標準モジュール
├── macros/
│   └── AutoExec.txt                # [simulated_access_export] SaveAsText (acMacro) 模擬
├── data/
│   ├── Employees.xlsx              # [data_source] 従業員データ（Excelテーブル: EmployeesTable）
│   └── AttendanceRecords.xlsx      # [data_source] 勤怠データ（Excelテーブル: AttendanceRecordsTable）
├── schema/
│   ├── tables.json                 # [custom_auxiliary] テーブル・カラム定義
│   ├── relationships.json          # [custom_auxiliary] 論理リレーションシップ定義
│   └── excel-links.json            # [custom_auxiliary] Excelリンク接続・書き込み制約定義
├── tests/
│   └── expected-behavior.json      # [evaluation_ground_truth] 評価用の期待動作定義
├── manifest.json                   # [custom_auxiliary] 構成ファイル一覧・分類メタデータ
└── README.md                       # [custom_auxiliary] 本ドキュメント
```

---

## 3. データ構造と Excel ファイル仕様

データソースとして 2 つの Excel ブックを使用し、それぞれに正式な Excel テーブルを設定しています。外部サービス（Excel Online (Business) 等）や Access の両方から安全に扱えるよう、日付・時刻も含めて複雑な型を避け、すべてテキスト形式（書式 `@` / ISO-8601 文字列）で統一しています。

### 3.1 `data/Employees.xlsx`

- ワークシート名: `Employees`
- Excel テーブル名: `EmployeesTable`
- 想定 Access テーブル名: `Employees`

| カラム名 | Access 型 | Excel 書式 | 必須 | 説明・値域 |
| :--- | :--- | :--- | :---: | :--- |
| `EmployeeID` | Short Text (10) | `@` | ◯ | 主キー（例: `EMP001`） |
| `EmployeeName` | Short Text (100) | `@` | ◯ | 従業員氏名（架空データ） |
| `Department` | Short Text (50) | `@` | ◯ | 所属部署 |
| `Status` | Short Text (20) | `@` | ◯ | `Active` / `Inactive` |

### 3.2 `data/AttendanceRecords.xlsx`

- ワークシート名: `AttendanceRecords`
- Excel テーブル名: `AttendanceRecordsTable`
- 想定 Access テーブル名: `AttendanceRecords`

| カラム名 | Access 型 | Excel 書式 | 必須 | 説明・値域 |
| :--- | :--- | :--- | :---: | :--- |
| `AttendanceID` | Short Text (20) | `@` | ◯ | 主キー（`ATT-YYYYMMDD-NNN` 形式） |
| `EmployeeID` | Short Text (10) | `@` | ◯ | 外部キー（`Employees.EmployeeID`） |
| `WorkDate` | Short Text (10) | `@` | ◯ | 勤務日（`YYYY-MM-DD`） |
| `ClockIn` | Short Text (19) | `@` | ◯ | 出勤日時（`YYYY-MM-DD HH:MM:SS`） |
| `ClockOut` | Short Text (19) | `@` | - | 退勤日時（`YYYY-MM-DD HH:MM:SS`、未退勤時は空） |
| `Status` | Short Text (20) | `@` | ◯ | `Working`（出勤中・未退勤） / `Completed`（退勤済） |

---

## 4. Access における Excel 接続上の制約

実際の Microsoft Access（ACE Excel ISAM ドライバ）には、**リンクされた Excel テーブルに対して DAO 経由での `UPDATE` / `DELETE` ができない制約（Microsoft KB 904953）** があります。

本構成ではこの実務上の制約を反映し、以下の設計としています（`schema/excel-links.json` および `modules/modAttendance.bas` を参照）。

- **参照系（フォーム・レポート・クエリ・入力検証）**: リンクテーブル `Employees` / `AttendanceRecords` に対して DAO およびドメイン集計関数（`DLookup`, `DCount`, `DMax`）で参照
- **更新系（出勤打刻の `INSERT` / 退勤打刻の `UPDATE`）**: `modAttendance.bas` 内で遅延バインディングの `ADODB.Connection` / `ADODB.Command`（`Microsoft.ACE.OLEDB.12.0`、`Extended Properties="Excel 12.0 Xml;HDR=YES;READONLY=FALSE;"`）を使用し、`AttendanceRecords.xlsx` の `[AttendanceRecords$]` に対してパラメータ化 SQL で書き込み
- **参照整合性**: Access は Excel リンクテーブルに物理的な参照整合性を設定できないため、`modValidation.ValidateEmployeeActive` によるアプリケーション検証で整合性を担保

---

## 5. 実機の `Application.SaveAsText` 出力との差異・既知の制約

本ファイル群は Microsoft Access を実行せずに生成しているため、実機のエクスポート出力とは以下の差異があります。

1. **文字コード・改行コード**: 実機の `Application.SaveAsText` は環境により UTF-16LE やローカルコードページ（Shift_JIS）で出力されますが、Git 管理とツール解析の可読性を優先して UTF-8（LF）で統一しています。
2. **バイナリ／チェックサム項目**: フォーム・レポート定義の `Checksum =0` はプレースホルダであり、実機固有のプリンタ設定バイナリ（`PrtMip`, `PrtDevMode`, `PrtDevNames`）やテーマバイナリは省略しています。そのため `Application.LoadFromText` での無修正インポートは保証しません。
3. **クエリ形式**: `Application.SaveAsText acQuery` は内部メタデータ付きのテキストブロックを出力しますが、解析しやすさを優先し標準構成どおり純粋な Access SQL（`.sql`）形式で出力しています。
4. **実行検証の未実施**: Microsoft Access / ACE ドライバ環境がないため、VBA のコンパイル・実行、フォーム起動、ADODB 経由の Excel 書き込みは実機で検証していません。
