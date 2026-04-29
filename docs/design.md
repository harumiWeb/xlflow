## 優先度高

### 1. `xlflow test`：VBAテスト実行ハーネス

最優先です。

```bash
xlflow test
xlflow test --filter TestCreateReport
```

VBA側に以下のような規約を作ります。

```vb
Public Sub TestCreateReport()
    AssertEquals 10, Sheets("Result").Range("A1").Value
End Sub
```

最低限ほしい機能はこれです。

- `Test*` / `*_Test` プロシージャを検出
- Excelを起動して実行
- 成功/失敗/エラーをCLIに出す
- JSON出力対応
- 終了コードでCI/AIエージェントが判断できる

```bash
xlflow test --json
```

これはAIエージェントにとっての生命線です。

---

### 2. `xlflow run`：マクロ実行ハーネス

任意のSubをCLIから実行できる仕組みです。

```bash
xlflow run Module1.CreateReport
xlflow run Sheet1.Button_Click
```

欲しい機能は以下です。

- 指定マクロの実行
- 引数指定
- 実行時間表示
- エラー時にVBAの行番号・モジュール名・Err.Number・Err.Descriptionを表示
- 実行前後でブックを保存するか選べる

```bash
xlflow run Report.Generate --input fixtures/sample.xlsx --save-as out/result.xlsx
```

---

### 3. `xlflow diff`：Excel/VBA差分ハーネス

AIが変更した結果を人間とAIが確認しやすくする機能です。

```bash
xlflow diff before.xlsm after.xlsm
```

見るべき差分は2種類あります。

#### VBAコード差分

```text
Module1.bas
- Range("A1").Value = "old"
+ Range("A1").Value = "new"
```

#### ワークブック状態差分

```text
Sheet: Report
A1: "" -> "売上レポート"
B2: 100 -> 120
Shape "Button1": caption changed
```

AIエージェントには「コード差分」だけでは不十分です。Excelは最終成果物がブック状態なので、セル・シート・名前定義・図形・印刷範囲あたりの差分があると強いです。

---

### 4. `xlflow lint`：VBA静的チェック

VBAは実行時まで壊れていることに気づきにくいので、軽量lintがあるとAIに効きます。

チェック例：

- `Option Explicit` がない
- 未使用変数
- 暗黙の `ActiveSheet` / `Selection` / `Activate`
- `On Error Resume Next` の濫用
- `Integer` 使用を `Long` 推奨
- グローバル状態の多用
- `Application.ScreenUpdating` を戻していない
- `Application.DisplayAlerts` を戻していない
- `Workbook_Open` などイベント系の危険変更

特にAI向けには、単なる警告ではなく修正指針があると良いです。

```text
Avoid ActiveSheet. Use explicit worksheet reference:
Set ws = ThisWorkbook.Worksheets("Report")
```

---

### 5. `xlflow doctor`：実行環境診断

Windows + Excel + COM + セキュリティ設定は壊れやすいので必須級です。

```bash
xlflow doctor
```

確認項目：

- Excelがインストールされているか
- COM起動できるか
- VBAプロジェクトへのアクセスが許可されているか
- マクロ実行ポリシー
- 一時ディレクトリ書き込み
- ブックが保護されていないか
- 既存のExcelプロセスが残っていないか

AIエージェントが「コードが悪い」のか「環境が悪い」のか判定できるようになります。

---

## あるとかなり強い

### 6. `xlflow trace`：実行ログ/イベント記録

VBAに簡易ロガーを差し込んで、AIが実行過程を見られるようにします。

```vb
Call XlflowLog("start GenerateReport")
Call XlflowLog("rowCount=" & rowCount)
```

CLI側：

```bash
xlflow run Report.Generate --trace
```

出力：

```text
[00:00.120] start GenerateReport
[00:00.214] rowCount=128
[00:00.420] created sheet: Report
```

VBAはデバッグ情報が取りづらいので、これはかなり効きます。

---

## 結論

xlflowで本当に価値が出るのは、VBAを書き出せることよりも、

```text
AIが変更する
↓
Excel上で実行する
↓
結果を機械的に検証する
↓
失敗理由をCLIで読める
↓
再修正する
```

このループを作れることです。
