# GUI操作と人間参加型の機能のデザイン案

## 1. GUIを「敵」として潰すのではなく、明示的な境界にする

VBAにはどうしても以下があります。

```vb
Application.GetOpenFilename
Application.FileDialog
MsgBox
InputBox
UserForm.Show
```

これらはAIエージェントの自動実行と相性が悪いです。

なので xlflow 側では、

> GUI操作が発生する箇所を検出し、
> 「ここは人間の操作が必要なチェックポイントです」と明示する

という扱いにするのが良いです。

たとえば `xlflow lint` や `xlflow doctor` で以下を検出します。

```txt
GUI boundary detected:
  - Main.bas:24 Application.GetOpenFilename
  - Main.bas:41 MsgBox
  - FormImport.frm UserForm.Show

Suggestion:
  This macro requires human interaction.
  Use --interactive or refactor GUI calls behind an adapter.
```

これは非常にAIエージェント向けです。
AIが「なぜ止まったのか分からない」状態を避けられます。

---

## 2. `run` に interactive / headless の2モードを作る

現状の `xlflow run` は、おそらく「AIが完全にCLIから実行する」前提が強いと思います。

ここに明示的に2モードを作ると良いです。

### headless mode

CI / AIエージェント向け。

```bash
xlflow run Main.ImportData --headless
```

この場合、GUI系APIを使っていたら即失敗でよいです。

```txt
Error: GUI operation is not allowed in headless mode.

Detected:
  Application.GetOpenFilename at Main.bas:24

Use:
  xlflow run Main.ImportData --interactive
or refactor file selection into parameters.
```

### interactive mode

人間共同作業向け。

```bash
xlflow run Main.ImportData --interactive
```

この場合は、ブックを開いたままExcel上でGUI操作を人間が行えるようにします。

```txt
Running macro in interactive mode...
Excel may show dialogs or message boxes.
Please complete the operation in Excel.

Waiting for macro completion...
```

これはかなり現実的です。

AIエージェントはCLIから `xlflow run --interactive` を実行し、人間がExcel上のファイル選択やMsgBox操作を担当する。
その後、xlflow が結果ログ・エラー・変更差分を回収する、という流れです。

---

## 3. 「人間参加型セッション」を明示的に作る

より踏み込むなら、こういうコマンドがあると強いです。

```bash
xlflow session start book.xlsm
```

または

```bash
xlflow attach book.xlsm
```

意味としては、

> すでに人間が開いているExcelブックにxlflowが接続する

です。

開発フローはこうなります。

```txt
1. 人間が Excel で対象ブックを開く
2. xlflow attach book.xlsm
3. AI が pull / lint / patch / push を行う
4. GUI操作が必要な macro は xlflow run --interactive で実行
5. 人間がファイル選択・MsgBox・UserForm操作をする
6. xlflow が実行結果・ログ・差分を取得
7. AI が次の修正を行う
```

これはかなり良い設計だと思います。

Excel/VBAの現実に合っています。
完全自動化にこだわるより、**AIが得意な編集・検査・差分管理・実行支援と、人間が得意なGUI判断を分担する** 方が実用性が高いです。

---

# 実装アイデア

## A. GUI APIの静的検出

まずは `lint` にルールを追加するのがよいです。

検出対象例：

```vb
MsgBox
InputBox
Application.GetOpenFilename
Application.GetSaveAsFilename
Application.FileDialog
UserForm.Show
DoEvents
Shell
CreateObject("WScript.Shell").Popup
```

ルール名は例えば：

```txt
XLW001: GUI interaction detected
XLW002: File picker detected
XLW003: Modal dialog detected
XLW004: UserForm detected
```

重要なのは、単に禁止するのではなく分類することです。

```txt
kind: file_picker
kind: modal_dialog
kind: user_form
kind: external_process
```

これによりAIエージェントが判断しやすくなります。

---

## B. `--headless` ではGUIを事前に拒否

AIエージェントが自動実行するときは、GUIが出た瞬間に詰みます。

なので `xlflow run --headless` は事前解析で落としてよいです。

```bash
xlflow run Main.ImportData --headless
```

```txt
Cannot run in headless mode because this macro uses GUI APIs.

- Application.GetOpenFilename
- MsgBox

Use --interactive if a human can operate Excel.
```

これは安全です。

---

## C. `--interactive` ではタイムアウト付きで待つ

```bash
xlflow run Main.ImportData --interactive --timeout 300
```

実行中にMsgBoxやファイル選択が出たら、人間が操作します。

タイムアウトしたら：

```txt
Macro did not complete within 300 seconds.
Possible causes:
  - File picker is still open
  - MsgBox is waiting for user input
  - UserForm is open
  - Macro is stuck in a loop
```

このメッセージが出るだけでも、AIエージェントにとってはかなり扱いやすくなります。

---

## D. GUI部分を「Adapter化」するリファクタ支援

これは xlflow の価値を大きく上げると思います。

たとえば元コードがこうだとします。

```vb
Sub ImportData()
    Dim path As Variant
    path = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx")

    If path = False Then Exit Sub

    Call ImportFromPath(CStr(path))
End Sub
```

AIエージェントには、こういう形にリファクタさせます。

```vb
Sub ImportData()
    Dim path As String
    path = PickImportFilePath()

    If path = "" Then Exit Sub

    Call ImportFromPath(path)
End Sub

Function PickImportFilePath() As String
    Dim selected As Variant
    selected = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx")

    If selected = False Then
        PickImportFilePath = ""
    Else
        PickImportFilePath = CStr(selected)
    End If
End Function
```

さらにテスト可能にするなら：

```vb
Sub ImportDataFromPath(ByVal path As String)
    If path = "" Then Exit Sub
    Call ImportFromPath(path)
End Sub
```

こうすると、AIは `ImportDataFromPath` をheadlessで実行できます。

```bash
xlflow run Main.ImportDataFromPath --args "C:\tmp\sample.xlsx"
```

一方、人間向けGUI入口は残せます。

```vb
Sub ImportData()
    Dim path As String
    path = PickImportFilePath()
    Call ImportDataFromPath(path)
End Sub
```

この設計はかなりおすすめです。

要するに、

```txt
GUI入口:
  ImportData

自動実行可能な本体:
  ImportDataFromPath(path)
```

に分ける。

xlflow のskillにも、このパターンを強く書くとよいです。

---

# xlflowとして用意するとよいコマンド案

## `xlflow inspect-gui`

```bash
xlflow inspect-gui
```

出力例：

```txt
GUI usage report

Main.bas
  12: Application.GetOpenFilename  [file_picker]
  28: MsgBox                       [modal_dialog]

Recommended refactor:
  - Extract file selection into PickImportFilePath()
  - Extract core logic into ImportDataFromPath(path)
```

AIエージェント向けにJSON出力もあるとよいです。

```bash
xlflow inspect-gui --json
```

```json
{
  "gui_boundaries": [
    {
      "file": "src/vba/Main.bas",
      "line": 12,
      "kind": "file_picker",
      "symbol": "Application.GetOpenFilename",
      "severity": "interactive-only"
    },
    {
      "file": "src/vba/Main.bas",
      "line": 28,
      "kind": "modal_dialog",
      "symbol": "MsgBox",
      "severity": "interactive-only"
    }
  ]
}
```

---

## `xlflow run --interactive`

```bash
xlflow run Main.ImportData --interactive
```

これは必須級です。

完全自動ではなく、人間がExcel上で操作する前提の実行モードです。

---

## `xlflow run --allow-gui`

`--interactive` と近いですが、より明示的にしてもよいです。

```bash
xlflow run Main.ImportData --allow-gui
```

個人的には `--interactive` の方が直感的です。

---

## `xlflow attach`

```bash
xlflow attach book.xlsm
```

または

```bash
xlflow attach --active
```

人間が開いているExcelインスタンスに接続できると、かなり実用的です。

理想は以下です。

```bash
xlflow attach --active
xlflow pull
xlflow push
xlflow run Main.ImportData --interactive
```

AIエージェントにとっても分かりやすいです。

---

# 「AIと人間が共同開発する」設計はありか？

かなりありです。

むしろExcel/VBAにおいては、これが一番現実的です。

WebアプリやCLIツールならheadless自動化しやすいですが、Excelは以下の制約があります。

```txt
- Excelアプリ本体がGUI前提
- COM操作が不安定になりやすい
- ファイル選択ダイアログがOS依存
- MsgBoxで処理が止まる
- UserFormは自動操作が難しい
- 既存業務ブックはGUI前提の作りが多い
```

なので xlflow は、

> VBAを完全にCLI化するツール

ではなく、

> Excel/VBA開発をAIエージェントが扱える形に整えるハーネス

として位置づける方が自然です。

この文脈では、人間参加型は弱点ではなく、むしろ設計思想になります。

---

# 推奨する設計方針

私なら、xlflow のドキュメントにこういう方針を書きます。

```txt
xlflow supports two execution styles:

1. Headless execution
   For AI agents, tests, CI, and repeatable automation.
   GUI APIs are treated as errors.

2. Interactive execution
   For real Excel workflows that require user interaction.
   xlflow runs the macro while a human operates Excel dialogs, message boxes, or forms.
```

そしてVBA設計ガイドとして、こう書きます。

```txt
Recommended VBA structure:

- Keep GUI code thin.
- Extract business logic into parameterized procedures.
- Make core procedures runnable without dialogs.
- Use GUI procedures only as human-facing entry points.
```

具体例：

```vb
' Human-facing entry point
Sub ImportData()
    Dim path As String
    path = PickImportFilePath()
    If path = "" Then Exit Sub

    ImportDataFromPath path
End Sub

' Agent/test-friendly entry point
Sub ImportDataFromPath(ByVal path As String)
    ' Core logic here
End Sub
```

この設計を xlflow skill にも入れるべきです。

---

# さらに強くするなら「対話チェックポイント」機能

面白い案として、`xlflow pause` 的な概念もあります。

VBA側にこういう関数を用意します。

```vb
Call XlflowCheckpoint("Please select the source file and click OK.")
```

実体はMsgBoxでもよいです。

```vb
Sub XlflowCheckpoint(ByVal message As String)
    MsgBox message, vbInformation, "xlflow checkpoint"
End Sub
```

xlflow側はログにこう出します。

```txt
Checkpoint reached:
  Please select the source file and click OK.

Waiting for user...
```

ただし、これは最初から凝りすぎる必要はありません。
まずは `--interactive` とGUI検出だけで十分価値があります。

---

# 優先順位

実装するなら、この順番が良いです。

## 優先度A

```txt
1. GUI API検出ルールをlint/doctorに追加
2. runに --headless / --interactive を導入
3. GUI検出時のエラーメッセージを改善
4. skillに「GUI入口とコア処理を分ける」設計原則を書く
```

## 優先度B

```txt
5. xlflow inspect-gui を追加
6. --json 出力でAIエージェントがGUI境界を読めるようにする
7. attach --active で開いているブックに接続
```

## 優先度C

```txt
8. checkpoint API
9. GUI操作待ち状態の推定
10. UserForm操作支援
```

---

# 最終的な方向性

かなり良い落とし所はこれです。

```txt
xlflowはGUIを完全自動化しようとしない。
代わりに、GUI境界を検出・明示・分離し、
必要な箇所だけ人間が操作できる interactive workflow を提供する。
```

この方針なら、Excel/VBAの現実に合っています。

特に良いのは、xlflow が以下の両方を扱えることです。

```txt
AI向け:
  headlessで再現可能な実行・検査・差分管理

人間向け:
  Excelを開いたままGUI操作込みで実行
```

この方向に進めると、xlflow は単なるVBA CLIツールではなく、**AIエージェント時代のExcel/VBA開発ハーネス** という立ち位置がかなり明確になります。
