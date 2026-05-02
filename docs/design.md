具体的には、**3層構成**で実現するのが現実的です。

```txt
1. VBA側ハーネス層
   実行時エラーを捕捉する

2. VBE/COM層
   コンパイル・選択中のコード位置・モジュール情報を取得する

3. Win32/UI Automation層
   VBEのモーダルダイアログを検出・読み取り・閉じる
```

xlflowの実装言語が Go/Rust なら、Windows専用部分として **COM + Win32 API + UI Automation** を使う設計になります。

---

# 1. 実行時エラーは「VBAラッパーモジュール」で捕捉する

これは一番素直です。

xlflowが一時的に以下のような標準モジュールをVBAプロジェクトに注入します。

```vb
Option Explicit

Public Sub __xlflow_run(ByVal entryPoint As String)
    On Error GoTo EH

    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Application.Run entryPoint

CleanUp:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

EH:
    __xlflow_write_error _
        "runtime", _
        Err.Number, _
        Err.Description, _
        Err.Source, _
        Erl

    Resume CleanUp
End Sub

Private Sub __xlflow_write_error( _
    ByVal kind As String, _
    ByVal number As Long, _
    ByVal description As String, _
    ByVal source As String, _
    ByVal lineNumber As Long)

    Dim path As String
    path = Environ$("TEMP") & "\xlflow-last-error.json"

    Dim f As Integer
    f = FreeFile

    Open path For Output As #f
    Print #f, "{"
    Print #f, "  ""kind"": """ & kind & ""","
    Print #f, "  ""number"": " & CStr(number) & ","
    Print #f, "  ""description"": """ & Replace(description, """", "\""") & ""","
    Print #f, "  ""source"": """ & Replace(source, """", "\""") & ""","
    Print #f, "  ""erl"": " & CStr(lineNumber)
    Print #f, "}"
    Close #f
End Sub
```

CLI側は、実行後にこのJSONを読むだけです。

```txt
%TEMP%\xlflow-last-error.json
```

出力例:

```txt
✖ VBA runtime error

Error     : 11
Message   : Division by zero
Source    : VBAProject
Line/Erl  : 120
```

## ただし行番号には「行番号注入」が必要

`Erl` は、VBAコードに行番号がないとほぼ `0` になります。

なので `xlflow run --diagnostic` では、実行前に一時的にこう変換します。

元コード:

```vb
Public Sub Main()
    Dim x As Long
    x = 1 / 0
End Sub
```

診断用コード:

```vb
Public Sub Main()
10  Dim x As Long
20  x = 1 / 0
End Sub
```

この状態なら `Erl = 20` が取れます。

xlflowとしては、元ソースを書き換えるのではなく、

```txt
source/*.bas
  ↓
一時ディレクトリにコピー
  ↓
行番号を注入
  ↓
Excelにpush
  ↓
実行
  ↓
終わったら元に戻す/破棄
```

が安全です。

---

# 2. コンパイルエラーは「VBE COM操作」で検出する

スクショのようなエラーはこれです。

```txt
コンパイル エラー:
メソッドまたはデータ メンバーが見つかりません。
```

これは実行時エラーではなく、**VBAプロジェクトのコンパイルエラー**です。

そのため `On Error GoTo` では捕まえられません。
対策は、実行前にVBEのコンパイルを走らせます。

## 技術的には VBE のメニューコマンドを実行する

Excel COMからVBEにアクセスして、VBEの「Debug > Compile VBAProject」を実行します。

概念的にはこうです。

```txt
Excel.Application
  └ VBE
      └ CommandBars("Debug")
          └ Controls("Compile VBAProject")
```

疑似コード:

```go
excel := com.CreateObject("Excel.Application")
vbe := excel.Get("VBE")

commandBars := vbe.Get("CommandBars")
debugBar := commandBars.Call("Item", "Debug")
compileButton := debugBar.Get("Controls").Call("Item", "Compile VBAProject")

compileButton.Call("Execute")
```

これでコンパイルが走ります。

コンパイル成功なら何も起きずに戻る。
コンパイル失敗なら、VBEがスクショのようなダイアログを出します。

---

# 3. コンパイルエラー箇所は「VBEの選択状態」から取る

コンパイルエラーが起きると、VBE上では問題箇所が選択されます。

スクショならここです。

```vb
.DisplayGridlines
```

つまり、ダイアログを閉じた後にVBEの `ActiveCodePane` を見ると、

```txt
どのモジュールか
何行目か
何列目か
どのトークンが選択されているか
```

を取れる可能性があります。

概念的にはこうです。

```go
activeCodePane := vbe.Get("ActiveCodePane")
codeModule := activeCodePane.Get("CodeModule")

var startLine, startColumn, endLine, endColumn int
activeCodePane.Call(
    "GetSelection",
    &startLine,
    &startColumn,
    &endLine,
    &endColumn,
)

moduleName := codeModule.Get("Name")
lineText := codeModule.Call("Lines", startLine, 1)
```

CLI出力:

```txt
✖ VBA compile error

Module : WeatherDashboard
Line   : 86
Column : 10
Token  : DisplayGridlines

Code:
  .DisplayGridlines = False
   ^^^^^^^^^^^^^^^^
```

ここまで取れれば、AIエージェントには十分使いやすいです。

---

# 4. ダイアログ本文は「UI Automation」か「Win32 API」で読む

一番厄介なのが、ダイアログの本文です。

```txt
コンパイル エラー:
メソッドまたはデータ メンバーが見つかりません。
```

これはExcel COMのエラーとして綺麗に返ってくるとは限りません。
そのため、Windows側からダイアログを読む必要があります。

候補は2つです。

---

## 方法: Win32 API

使うAPIはこのあたりです。

```txt
EnumWindows
FindWindow
FindWindowEx
GetWindowTextW
GetClassNameW
SendMessageW
PostMessageW
SetForegroundWindow
```

流れはこうです。

```txt
1. Excel/VBEプロセスIDを把握する
2. そのプロセスに属するトップレベルウィンドウを列挙する
3. タイトルが "Microsoft Visual Basic for Applications" のダイアログを探す
4. 子ウィンドウを列挙する
5. Staticコントロールのテキストを読む
6. OKボタンに BM_CLICK を送る
```

疑似コード:

```go
EnumWindows(func(hwnd HWND) bool {
    pid := GetWindowThreadProcessId(hwnd)

    if pid != excelPid {
        return true
    }

    title := GetWindowText(hwnd)
    className := GetClassName(hwnd)

    if title == "Microsoft Visual Basic for Applications" {
        // 子コントロールを列挙
        EnumChildWindows(hwnd, func(child HWND) bool {
            text := GetWindowText(child)
            cls := GetClassName(child)

            // Static: メッセージ本文
            // Button: OK / ヘルプ
            collect(cls, text)
            return true
        })

        // OKボタンを探してクリック
        okButton := FindChildButton(hwnd, "OK")
        SendMessage(okButton, BM_CLICK, 0, 0)

        return false
    }

    return true
})
```

取れる情報例:

```json
{
  "title": "Microsoft Visual Basic for Applications",
  "text": ["コンパイル エラー:", "メソッドまたはデータ メンバーが見つかりません。"],
  "buttons": ["OK", "ヘルプ"]
}
```

---

# 5. 「GUI完全抑制」は実際には watcher 方式になる

完全にダイアログを出さない、というより、実装上はこうなります。

```txt
xlflow run
  ↓
Excel起動
  ↓
別スレッドでダイアログ監視開始
  ↓
compile実行
  ↓
ダイアログが出た瞬間に検出
  ↓
本文を読む
  ↓
OKを押して閉じる
  ↓
VBEの選択位置を読む
  ↓
CLIに出す
```

つまり、ユーザーから見ると一瞬で閉じるため「GUIが出ていない」ように見えます。
ただし技術的には、**出たダイアログを即座に捕まえて閉じる**方式です。

これはかなり現実的です。

---

# 6. 実行フロー全体

xlflowとしてはこういう実装がよいです。

```txt
xlflow run Main --diagnostic

1. Excelを非表示で起動
   Application.Visible = False
   DisplayAlerts = False
   EnableEvents = False

2. 対象ブックを開く

3. ソースをpush
   必要なら行番号注入済みソースをpush

4. VBEダイアログ監視スレッドを開始

5. Compile VBAProject を実行

6. コンパイルエラーが出たら
   - ダイアログ本文を読む
   - OKを押して閉じる
   - ActiveCodePaneから行・列・モジュール名を読む
   - CLIに出して終了

7. コンパイル成功なら
   __xlflow_run 経由で Application.Run

8. 実行時エラーが出たら
   - VBAハーネスがJSONに保存
   - CLIがJSONを読む
   - CLIに出す

9. Excel状態を復元して終了
```

図にするとこうです。

```txt
┌─────────────┐
│ xlflow CLI  │
└──────┬──────┘
       │
       v
┌──────────────────────┐
│ Excel COM Automation  │
└──────┬───────────────┘
       │
       ├─ push VBA source
       │
       ├─ inject __xlflow_run
       │
       ├─ run VBE compile
       │
       │
       ├──────────────┐
       │              v
       │      ┌────────────────┐
       │      │ Dialog Watcher │
       │      │ Win32 / UIA    │
       │      └────────────────┘
       │
       ├─ read VBE selection
       │
       └─ Application.Run
```

---

# 7. Goで実装する場合の技術候補

Goなら、おそらくこの構成です。

## COM操作

候補:

```txt
github.com/go-ole/go-ole
github.com/go-ole/go-ole/oleutil
```

用途:

```txt
Excel.Application 操作
Workbook.Open
Application.Run
VBE操作
CommandBars操作
VBComponents操作
CodeModule操作
```

## Win32 API

候補:

```txt
golang.org/x/sys/windows
github.com/lxn/win
```

用途:

```txt
EnumWindows
EnumChildWindows
GetWindowTextW
GetClassNameW
SendMessageW
PostMessageW
GetWindowThreadProcessId
```

---

# 8. 一番堅いMVP

いきなり完全対応を狙うより、MVPはこれがよいです。

## MVP 1: runtime error CLI化

```txt
- __xlflow_run を注入
- Err.Number / Err.Description / Erl をJSON出力
- CLIで表示
```

これは比較的すぐ実装できます。

## MVP 2: compile-first

```txt
- VBEのCompileコマンドを実行
- 失敗時はダイアログ監視で閉じる
- ActiveCodePaneからモジュール・行・列を読む
```

これでスクショのようなエラーに対応できます。

## MVP 3: ダイアログ本文取得

```txt
- Win32 APIでVBEダイアログを列挙
- Staticテキストを読む
- OKボタンを押す
```

CLI出力が一気に実用的になります。

---

# 9. 実装上の注意点

## 1. 「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」が必要

VBEのコードモジュールを操作する場合、Excel側の設定で

```txt
VBAプロジェクト オブジェクト モデルへのアクセスを信頼する
```

が有効である必要があります。

xlflowの `doctor` でチェックすべきです。

```txt
xlflow doctor

✓ Excel installed
✓ COM automation available
✓ Trust access to VBA project object model enabled
```

## 2. 言語環境でダイアログ文言が変わる

日本語環境:

```txt
Microsoft Visual Basic for Applications
コンパイル エラー:
メソッドまたはデータ メンバーが見つかりません。
OK
ヘルプ
```

英語環境:

```txt
Microsoft Visual Basic for Applications
Compile error:
Method or data member not found
OK
Help
```

なので、メッセージ本文のパースは多言語対応を考える必要があります。

ただし、最低限は本文をそのままCLIに出すだけで十分です。

## 3. `DisplayAlerts = False` だけでは無理

これは重要です。

```vb
Application.DisplayAlerts = False
```

だけでは、VBEのコンパイルエラーダイアログや `MsgBox` は消えません。

したがって、以下を分けて考える必要があります。

```txt
Excel標準警告:
  DisplayAlerts = False

VBA実行時エラー:
  On Error GoTo + Err

VBEコンパイルエラー:
  VBE Compile + Win32/UIA dialog watcher

ユーザーコードのMsgBox/FileDialog:
  lint検出 + headless-strictで禁止
```

---

# 10. CLI設計案

こういう出力にするとAIエージェントに効きます。

```bash
xlflow run Main --diagnostic
```

コンパイルエラー:

```txt
✖ VBA compile error

Workbook : weather.xlsm
Module   : WeatherDashboard
Line     : 86
Column   : 10
Token    : DisplayGridlines

Message:
  コンパイル エラー:
  メソッドまたはデータ メンバーが見つかりません。

Code:
  With ws
      .DisplayGridlines = False
       ^^^^^^^^^^^^^^^^
  End With

Hint:
  DisplayGridlines is a Window property, not a Worksheet property.
  Try:
    ws.Activate
    ActiveWindow.DisplayGridlines = False
```

実行時エラー:

```txt
✖ VBA runtime error

Workbook : weather.xlsm
Entry    : Main
Error    : 1004
Message  : Application-defined or object-defined error
Source   : VBAProject
Line     : 230

Code:
  .Range("A1").Value = data("title")
```

GUI依存検出:

```txt
✖ Headless run blocked

The macro contains GUI-dependent APIs.

Module : ImportDialog
Line   : 42
Code   : Application.FileDialog(msoFileDialogFilePicker).Show

Use:
  xlflow run Main --interactive
```

---

# 11. 私ならこう実装します

xlflowのロードマップとしては、この順番が一番安全です。

```txt
Phase 1:
  runtime error harness
  Err.Number / Err.Description / Erl JSON出力

Phase 2:
  diagnostic line-number injection
  実行時エラー行のCLI表示

Phase 3:
  compile-first
  VBE Compile実行

Phase 4:
  VBE dialog watcher
  Win32 APIでダイアログ本文取得・OK自動クリック

Phase 5:
  ActiveCodePane解析
  Module / Line / Column / Token 表示

Phase 6:
  GUI-dependent API lint
  MsgBox / InputBox / FileDialog / UserForm.Show 検出

Phase 7:
  --headless-strict / --interactive 分岐
```

この構成なら、スクショのようなコンパイルエラーも、VBA実行中のランタイムエラーも、AIエージェントが読めるCLI出力に変換できます。

技術的な核はこの4つです。

```txt
Excel COM
VBE COM
Win32 API or UI Automation
VBA実行ハーネス
```
