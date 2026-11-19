# 現行 Microsoft Access 勤怠管理システム 構成ファイル (`access-source`)

## 1. 概要

本ディレクトリは、社内で長年改修を重ねて運用されてきた従来型 Microsoft Access 勤怠管理システム（`AttendanceManagement` / Legacy Ver 3.4）を想定し、`Application.SaveAsText`・VBE モジュールエクスポート・クエリ SQL 抽出によって取得した構成ファイル群を疑似的に再現したものです。

Microsoft Access の実行環境および `.accdb` ファイルは含みません。詳細な画面挙動、コード定義、計算ロジック、権限ルールについては、各フォーム定義・マクロ・SQL クエリ・VBA 標準モジュールの実体ファイルを参照してください。

---

## 2. ディレクトリ構成とファイル分類

ファイルの種別は [manifest.json](manifest.json) にて `simulated_access_export`（Access 実エクスポート模擬）、`data_source`（実 Excel ブック）、`custom_auxiliary`（補助スキーマ・台帳）、`evaluation_ground_truth`（評価基準）に分類しています。

```text
access-source/
├── forms/                          # [simulated_access_export] SaveAsText (acForm)
│   ├── frmMenu.txt                 # スイッチボード型メインメニュー
│   ├── frmMain.txt                 # タイムカード打刻画面
│   ├── frmEmployees.txt            # 従業員マスタ一覧
│   ├── frmAttendance.txt           # 月次出勤簿（親フォーム）
│   ├── sfrmAttendanceDetail.txt    # 月次出勤簿明細（サブフォーム）
│   ├── frmCorrectionRequest.txt    # 打刻修正申請・上長承認
│   └── frmMonthlyClose.txt         # 月次締め管理・CSV出力
├── reports/                        # [simulated_access_export] SaveAsText (acReport)
│   ├── rptAttendanceHistory.txt    # 従業員別 月次出勤簿レポート
│   └── rptMonthlyDepartment.txt    # 部署別 月次残業サマリーレポート
├── queries/                        # [simulated_access_export] Access SQL / Crosstab SQL
│   ├── qryEmployees.sql
│   ├── qryOpenAttendance.sql
│   ├── qryAttendanceBase.sql
│   ├── qryAttendanceCalc.sql
│   ├── qryAttendanceHistory.sql
│   ├── qryMonthlySummary.sql
│   ├── qryMonthlyCrosstab.sql
│   ├── qryPendingCorrections.sql
│   └── qryAttendance_old_bk.sql    # 旧版残置オブジェクト（詳細要確認）
├── modules/                        # [simulated_access_export] VBA 標準モジュール (.bas)
│   ├── modGlobals.bas
│   ├── modValidation.bas
│   ├── modCalc.bas
│   ├── modAttendance.bas
│   └── modLegacyExport.bas
├── macros/                         # [simulated_access_export] SaveAsText (acMacro)
│   ├── AutoExec.txt
│   └── mcrDailyOps.txt
├── data/                           # [data_source] 正式な Excel テーブル付き .xlsx
│   ├── Employees.xlsx              # EmployeesTable
│   ├── WorkPatterns.xlsx           # WorkPatternsTable
│   ├── AttendanceRecords.xlsx      # AttendanceRecordsTable
│   ├── CorrectionRequests.xlsx     # CorrectionRequestsTable
│   └── MonthlyClosings.xlsx        # MonthlyClosingsTable
├── schema/                         # [custom_auxiliary] 機械可読スキーマ・リンク設定
│   ├── tables.json
│   ├── relationships.json
│   └── excel-links.json
├── tests/
│   └── expected-behavior.json      # [evaluation_ground_truth] 評価用期待動作定義
├── manifest.json                   # [custom_auxiliary] 構成ファイル台帳
└── README.md                       # [custom_auxiliary] 本ドキュメント
```

---

## 3. Excel データソース仕様

全 5 ブックについて、外部連携サービスおよび Access の双方から型不整合なく扱えるよう、正式な Excel テーブルを設定し、全セルをテキスト書式（`@`）・ISO-8601 文字列で統一しています（数式・結合セル不使用）。各カラム仕様は [schema/tables.json](schema/tables.json) を参照してください。

| ブック | シート名 | Excel テーブル名 | 主キー |
| :--- | :--- | :--- | :--- |
| `data/Employees.xlsx` | `Employees` | `EmployeesTable` | `EmployeeID` |
| `data/WorkPatterns.xlsx` | `WorkPatterns` | `WorkPatternsTable` | `PatternCode` |
| `data/AttendanceRecords.xlsx` | `AttendanceRecords` | `AttendanceRecordsTable` | `AttendanceID` |
| `data/CorrectionRequests.xlsx` | `CorrectionRequests` | `CorrectionRequestsTable` | `RequestID` |
| `data/MonthlyClosings.xlsx` | `MonthlyClosings` | `MonthlyClosingsTable` | `TargetYM` |

---

## 4. Access 固有の構成・制約に関する注記

1. **Excel リンクテーブル更新制約（Microsoft KB 904953）**:
   Access の ACE Excel ISAM ではリンクされた Excel テーブルへの DAO `UPDATE` / `DELETE` が不可のため、参照系はリンクテーブル・クエリ・ドメイン集計関数を使用し、更新系（`INSERT` / `UPDATE`）は `modGlobals.OpenExcelWriteConnection` 経由の ADODB（`Microsoft.ACE.OLEDB.12.0`、`READONLY=FALSE`）を使用しています（[schema/excel-links.json](schema/excel-links.json) 参照）。
2. **ロジックの分散構成**:
   Access 特有の設計として、起動マクロ・サブマクロ・フォーム埋め込みマクロ・計算コントロール式・フォーム参照パラメータ付きクエリ・VBA ユーザー定義関数を直接呼び出す SQL クエリ・クロス集計クエリ（`TRANSFORM ... PIVOT`）・親子サブフォーム連動（`LinkMasterFields` / `LinkChildFields`）・`TempVars` / グローバル変数 / 隠しコントロールが混在しています。
3. **実機 `Application.SaveAsText` 出力との差異**:
   - Git・静的解析の可読性のため文字コードは UTF-8 (LF) に統一しています（実機は UTF-16LE またはローカルコードページ）。
   - `Checksum =0` はプレースホルダであり、`PrtMip` / `PrtDevMode` 等のバイナリブロックは省略しています。
   - クエリは SQL 解析の容易性を考慮し `.sql` 形式で配置しています。
   - Microsoft Access 実機環境がないため、VBA コンパイル・フォーム起動・ADODB 書込のランタイム動作は未検証です。
