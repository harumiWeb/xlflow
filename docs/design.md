# xlflow の次の一手

## 次に「失敗理由の見える化」を強化する

レビューの中で一番価値が高い改善要望はここです。

> 450 系や型の問題が出たときに、どの行でどの代入が原因かをもう少し直接に示してほしい

これは xlflow の差別化要素になり得ます。

Excel VBA の自動開発で一番つらいのは、実行時エラーが出ても、

- どのモジュールか
- どのプロシージャか
- どの行か
- 直前に何をしていたか
- Excel COM 側の失敗か
- VBA の型エラーか
- ファイル IO か
- 外部 API か

が分かりにくいことです。

### 改善案

`xlflow run` の失敗時に、単なるエラー表示ではなく、診断レポート形式にするのがよいです。

```txt
XLFLOW_RUN_FAILED

Error:
  VBA Runtime Error 450
  Wrong number of arguments or invalid property assignment

Location:
  Module: Main
  Procedure: BuildReport
  Line: 42

Likely cause:
  Object value was assigned without Set, or a property/method call is invalid.

Nearby code:
  40 | Dim ws As Worksheet
  41 | Dim result As Range
> 42 | result = FindTargetRange(ws)
  43 | result.Value = "OK"

Suggestion:
  If FindTargetRange returns an object, use:
    Set result = FindTargetRange(ws)
```

これができると、単なる VBA 実行ラッパーではなく、**VBA 専用の開発ハーネス**になります。

### 実装の現実性

完全な静的解析で原因特定するのは難しいですが、かなり実用的な近似はできます。

特に以下は検出しやすいです。

```vb
Dim ws As Worksheet
ws = ThisWorkbook.Worksheets("Sheet1")
```

これは `Set` が必要。

```vb
Dim rng As Range
rng = FindRange()
```

これも `FindRange()` の戻り値が `Range` なら `Set` が必要。

```vb
Function GetSheet() As Worksheet
    GetSheet = ThisWorkbook.Worksheets(1)
End Function
```

これも `Set GetSheet = ...` が必要。

まずは厳密な型推論ではなく、**VBAあるある実行時エラーのパターン検出**から始めるのが現実的です。

---

## `xlflow doctor` / `xlflow analyze` を強化する

レビューにある、

> 実行前の軽い型・参照チェック
> Set が必要な箇所
> 存在しないキー参照
> オブジェクト返却の受け方

これは `lint` とは別枠にした方がよいです。

`lint` はコードスタイル・危険構文チェック。
`analyze` は実行時エラー予防。
`doctor` は環境・ブック・Excel COM 状態確認。

という役割分担がきれいです。

### 例

```bash
xlflow lint
xlflow analyze
xlflow doctor
```

または、

```bash
xlflow check
```

でまとめて実行。

```txt
xlflow check

[lint]
  OK Option Explicit is present
  WARN Select/Activate usage found

[analyze]
  WARN Possible missing Set
    Module: Main
    Line: 42
    Code: result = FindTargetRange(ws)
    Reason: result is declared As Range

[doctor]
  OK Excel COM is available
  OK Trust access to VBA project object model is enabled
  OK Workbook exists
```

この `analyze` はかなり強いです。
VBA の AI 開発では、実行前に潰せるミスを潰すだけで往復回数が大きく減ります。

---

## trace は「導入手順」ではなく「状態切替コマンド」にする

レビューの、

> trace 周りは導入手順の見通しがやや弱い
> trace の有効化と無効化をワンコマンドで切り替える仕組み

これはその通りだと思います。

trace は便利でも、手順が面倒だと AIエージェントも人間も使いません。

### 改善案

```bash
xlflow trace enable
xlflow trace disable
xlflow trace status
xlflow trace clean
```

があるとかなり分かりやすいです。

さらに、実行時に一時的に有効化できると便利です。

```bash
xlflow run Main.Run --trace
```

この場合は、

1. trace helper を注入
2. 実行
3. trace log を回収
4. 必要なら元に戻す

までやる。

```txt
[XLFLOW] Trace enabled temporarily
[XLFLOW] Running Main.Run
[XLFLOW] Trace log saved: .xlflow/traces/2026-05-01-001.log
[XLFLOW] Trace helper reverted
[XLFLOW_DONE]
```

これなら常用コードに trace 支援コードを残す必要が減ります。

## テンプレート整備

レビューにある、

> workbook open と macro entrypoint の役割分担
> GUI を持つブックでも、実処理は Main.Run のような薄い入口と、引数を受ける core に分けるのが安定

これは xlflow の公式テンプレートにした方がよいです。

### 推奨構成

```vb
' Main.bas
Option Explicit

Public Sub Run()
    App.RunCore ThisWorkbook
End Sub
```

```vb
' App.bas
Option Explicit

Public Sub RunCore(ByVal wb As Workbook)
    ' 実処理
End Sub
```

```vb
' Ui.bas
Option Explicit

Public Sub RunFromButton()
    Main.Run
End Sub
```

```vb
' Workbook_Open
Private Sub Workbook_Open()
    ' 自動実行が必要な場合だけ薄く呼ぶ
    ' Main.Run
End Sub
```

この形にしておくと、

- xlflow headless run
- ボタン実行
- Workbook_Open
- 手動実行
- テスト用 entrypoint

を分けやすくなります
