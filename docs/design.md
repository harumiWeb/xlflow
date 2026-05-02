# xlflow inspect コマンドの設計メモ

方向性としては、

> VBA開発後に、AIエージェントが「ブックの状態・シート構成・セル出力」をCLIから確認できる最小機能

に限定するのが良いです。

## 実装するならコマンド名は `xlflow inspect` 系が自然

すでに `run` / `push` / `pull` / `lint` / `doctor` のような開発系コマンドがあるなら、セルやシートの確認は `inspect` にまとめるのが扱いやすいです。

```bash
xlflow inspect workbook
xlflow inspect sheets
xlflow inspect range "Sheet1!A1:F20"
xlflow inspect used-range "Sheet1"
xlflow inspect cell "Sheet1!A1"
```

ただ、サブコマンドが増えすぎるなら、最初はこの3つだけでも十分です。

```bash
xlflow inspect workbook
xlflow inspect sheets
xlflow inspect range "Sheet1!A1:F20"
```

## 最小MVP仕様

### 1. ブック概要

```bash
xlflow inspect workbook --book sample.xlsm --format json
```

出力例:

```json
{
  "path": "C:\\work\\sample.xlsm",
  "name": "sample.xlsm",
  "sheets": [
    {
      "name": "World News",
      "index": 1,
      "visible": true,
      "usedRange": "A1:F18"
    }
  ],
  "activeSheet": "World News"
}
```

これはAIにとってかなり重要です。
「どのシートがあるか」「出力先シート名が合っているか」「usedRangeがどこまであるか」を確認できます。

### 2. シート一覧

```bash
xlflow inspect sheets --book sample.xlsm --format json
```

出力例:

```json
[
  {
    "name": "World News",
    "index": 1,
    "visible": true,
    "usedRange": "A1:F18",
    "rowCount": 18,
    "columnCount": 6
  },
  {
    "name": "Config",
    "index": 2,
    "visible": false,
    "usedRange": "A1:B5",
    "rowCount": 5,
    "columnCount": 2
  }
]
```

AIエージェントが「シート作成に成功したか」「余計なシートを作っていないか」を確認できます。

### 3. セル範囲の値取得

```bash
xlflow inspect range --book sample.xlsm --range "World News!A1:F10" --format json
```

出力例:

```json
{
  "sheet": "World News",
  "range": "A1:F10",
  "values": [
    ["World News", null, null, null, null, null],
    ["Latest stories from NewsAPI", null, null, null, null, null],
    ["Updated: 2026-05-02 18:32:44 | Query: world | Showing: 13", null, null, null, null, null],
    [null, null, null, null, null, null],
    ["Published", "Source", "Image", "Title", "Description", "Link"],
    [
      "2026-05-01T09:28:51Z",
      "CNA",
      null,
      "Liverpool's Salah ruled out of Man Utd clash, says Slot",
      "...",
      "Open article"
    ]
  ]
}
```

これは最重要です。
AIにとって「実行結果の妥当性確認」ができるようになります。

## まずは値だけでいい

初期版では、以下は不要だと思います。

- 図形取得
- 画像取得
- グラフ取得
- セル色
- フォント
- 罫線
- 結合セル
- 数式依存関係
- 印刷範囲
- オートフィルター
- テーブル構造推定

ここまでやると exstruct に寄ってしまいます。

xlflow の inspect は、あくまで

> AIエージェントがVBA実行後の結果を確認するための軽量スナップショット

でよいです。

## ただし、ハイパーリンクと数式はあると嬉しい

軽量機能の範囲でも、次の2つは実用性が高いです。

```bash
xlflow inspect range "World News!A1:F10" --include formulas
xlflow inspect range "World News!A1:F10" --include hyperlinks
```

出力例:

```json
{
  "sheet": "World News",
  "range": "A1:F10",
  "cells": [
    {
      "address": "F6",
      "value": "Open article",
      "formula": null,
      "hyperlink": "https://example.com/article"
    }
  ]
}
```

ただし、毎回この形式だと重いので、デフォルトは2次元配列で十分です。

おすすめはこうです。

```bash
# デフォルト: 値だけ、軽量
xlflow inspect range "Sheet1!A1:F20"

# 詳細: セルごとのメタ情報
xlflow inspect range "Sheet1!A1:F20" --mode cells --include formulas,hyperlinks
```

## 出力形式は3種類あると強い

AIエージェント向けには JSON が最重要です。

```bash
--format json
```

人間向けには table。

```bash
--format table
```

LLM向けには markdown も便利です。

```bash
--format markdown
```

例:

```bash
xlflow inspect range "World News!A5:F8" --format markdown
```

```md
| Published            | Source | Image | Title                          | Description         | Link         |
| -------------------- | ------ | ----- | ------------------------------ | ------------------- | ------------ |
| 2026-05-01T09:28:51Z | CNA    |       | Liverpool's Salah ruled out... | May 1: Liverpool... | Open article |
```

個人的には、MVPでは `json` と `markdown` があれば十分です。
`table` は後からでよいです。

## 範囲指定の仕様

A1表記を基本にするのが良いです。

```bash
xlflow inspect range "Sheet1!A1:F20"
```

シート名に空白がある場合も対応。

```bash
xlflow inspect range "'World News'!A1:F20"
```

ただしCLIではクォートが面倒なので、別オプションもあると便利です。

```bash
xlflow inspect range --sheet "World News" --address "A1:F20"
```

AIエージェントにはこちらの方が安定します。

```bash
xlflow inspect range --sheet "World News" --address "A1:F20" --format json
```

## used range 取得はかなり重要

実行結果の全体を見るには、AIは範囲を知らないことが多いです。
そのため、`used-range` はあると非常に便利です。

```bash
xlflow inspect used-range --sheet "World News" --format markdown
```

ただし、巨大シートで爆発しないように制限が必要です。

```bash
xlflow inspect used-range --sheet "World News" --max-rows 50 --max-cols 20
```

出力には省略情報を入れるとよいです。

```json
{
  "sheet": "World News",
  "usedRange": "A1:F120",
  "returnedRange": "A1:F50",
  "truncated": true,
  "maxRows": 50,
  "maxCols": 20,
  "values": []
}
```

これはかなり大事です。
AIエージェントは巨大出力をそのまま食わせると壊れやすいので、**デフォルトで安全側に切る**べきです。

## おすすめのデフォルト制限

初期値はこのくらいが良いです。

```text
max rows: 100
max columns: 30
max cell text length: 500
max total cells: 3000
```

超えたらエラーではなく、基本は truncate でよいです。

```json
{
  "truncated": true,
  "warnings": [
    "Output was truncated: used range has 1200 rows x 42 columns, returned 100 rows x 30 columns."
  ]
}
```

AIにやさしいです。

## `run` と連携すると強い

最終的には、`run` の後に inspect しやすくする設計が良いです。

```bash
xlflow run Main.FetchNews
xlflow inspect used-range --sheet "World News" --format markdown
```

さらに将来的には、実行後に自動でサマリを出すオプションもありです。

```bash
xlflow run Main.FetchNews --inspect "World News!A1:F20"
```

ただし最初から入れる必要はありません。
まずは独立コマンドで十分です。

## 画像については「存在確認」だけでよい

NewsAPIの例では画像が重要に見えますが、inspectで画像そのものを取る必要はないです。

やるとしても、軽量にこれだけで十分です。

```bash
xlflow inspect objects --sheet "World News"
```

出力例:

```json
{
  "sheet": "World News",
  "pictures": [
    {
      "name": "Picture 1",
      "topLeftCell": "C6",
      "width": 180,
      "height": 95
    }
  ]
}
```

ただ、これはMVPから外してよいです。
まずはセル値・シート・usedRangeだけで十分価値があります。

## 仕様案としてはこれが一番きれいです

```bash
xlflow inspect workbook --book sample.xlsm --format json

xlflow inspect sheets --book sample.xlsm --format json

xlflow inspect range --book sample.xlsm \
  --sheet "World News" \
  --address "A1:F20" \
  --format json

xlflow inspect used-range --book sample.xlsm \
  --sheet "World News" \
  --max-rows 100 \
  --max-cols 30 \
  --format markdown
```

## JSON形式は2段階に分けるのがおすすめ

### デフォルト: matrix mode

```json
{
  "sheet": "World News",
  "range": "A1:F10",
  "mode": "matrix",
  "values": [["Published", "Source", "Image", "Title", "Description", "Link"]]
}
```

### 詳細: cells mode

```json
{
  "sheet": "World News",
  "range": "A1:F10",
  "mode": "cells",
  "cells": [
    {
      "address": "A5",
      "row": 5,
      "column": 1,
      "value": "Published",
      "formula": null,
      "numberFormat": "General",
      "hyperlink": null
    }
  ]
}
```

MVPは `matrix` のみでよいです。
`cells` は後から追加で十分です。

## 内部実装はCOMで十分

xlflowはWindows + Excel前提のCLIなので、ここはExcel COMでよいと思います。

- `Workbook.Worksheets`
- `Worksheet.UsedRange`
- `Range.Value2`
- `Range.Formula`
- `Range.Hyperlinks`
- `Worksheet.Visible`
- `Application.ActiveSheet`

この程度なら重くありません。
exstruct のような shape/chart/print-area 深掘りとはまったく別物です。

## 実装優先順位

おすすめはこの順番です。

### Phase 1: 最小inspect

```bash
xlflow inspect sheets
xlflow inspect range --sheet <name> --address <A1:B2>
xlflow inspect used-range --sheet <name>
```

取得対象:

```text
sheet name
sheet index
visible
used range address
cell values
```

### Phase 2: AI向け出力改善

```bash
--format json
--format markdown
--max-rows
--max-cols
--max-cell-chars
--null-as-empty
```

### Phase 3: 開発デバッグ向け

```bash
--include formulas
--include hyperlinks
--mode cells
```

### Phase 4: 軽量オブジェクト確認

```bash
xlflow inspect objects --sheet <name>
```

取得対象は picture/chart/shape の数と位置だけ。

## 命名で迷うなら `read` でもよい

`inspect` は少し広い名前です。
セル値取得に寄せるなら `read` も自然です。

```bash
xlflow read range --sheet "World News" --address "A1:F20"
xlflow read used-range --sheet "World News"
```

ただ、シート情報やブック情報も含めるなら `inspect` の方が良いです。

私ならこうします。

```bash
xlflow inspect sheets
xlflow inspect range
xlflow inspect used-range
```

## 結論

xlflowに入れるなら、**「軽量Excel状態観測コマンド」**として実装するのがよいです。

最初の仕様はこれで十分です。

```bash
xlflow inspect sheets --format json

xlflow inspect range \
  --sheet "World News" \
  --address "A1:F20" \
  --format json

xlflow inspect used-range \
  --sheet "World News" \
  --max-rows 100 \
  --max-cols 30 \
  --format markdown
```

この機能が入ると、AIエージェントはかなり自律的に回せます。

> VBAを書く → push → run → inspect → 結果を見て修正

というループが完成するので、xlflowの「AIエージェント向けExcel VBA開発ハーネス」としての価値が一段上がります。

結論から言うと、**xlflow の `inspect` 系はまず excelize 優先で実装するのが良い**と思います。
ただし、**開いているExcelアプリ上の最新状態を見たい機能だけはCOM**が必要です。

つまり、おすすめは二層構成です。

```txt
基本: excelize backend
必要時: COM backend
```

## 基本は excelize が向いています

今回やりたい簡易機能が、

```bash
xlflow inspect sheets
xlflow inspect range
xlflow inspect used-range
```

くらいなら、excelize で十分です。

xlflow の文脈では、`inspect` は「AIエージェントがファイルの状態を確認する」ための機能なので、**Excelを起動しなくても読める**ことにかなり価値があります。

excelize のメリットは大きいです。

```txt
- Excel COMを起動しないので速い
- GUIダイアログに巻き込まれにくい
- Windows以外でも動く可能性がある
- AIエージェント向けCLIとして安定しやすい
- 依存関係としてすでに入っているなら追加コストが小さい
- セル値、シート一覧、行データ、結合セル、数式、ハイパーリンク程度なら十分取れる
```

特に `inspect range` は excelize 向きです。

```bash
xlflow inspect range --sheet "World News" --address "A1:F20" --format json
```

これはファイルを直接読めばよく、Excel本体を操作する必要がありません。

## ただし COM が必要になる場面もあります

COMが必要なのは、**Excelアプリ上の状態を読む必要がある場合**です。

たとえば以下です。

```txt
- ユーザーがExcelを開いたまま編集していて、まだ保存していない内容を読みたい
- VBA実行直後のブック状態を、保存せずにそのまま読みたい
- マクロ実行で作られた画像、図形、グラフ、ボタンなどの配置を確認したい
- 数式の計算結果をExcelエンジンで再計算した後に読みたい
- UsedRangeをExcel本体の認識と同じように取得したい
- 非表示/VeryHidden、ActiveSheet、SelectionなどExcelアプリ状態を確認したい
```

excelize は基本的に **保存済みファイルのOpenXMLを読む** ものです。
そのため、Excel上で変更されたが保存されていない内容は読めません。

ここはかなり重要です。

## xlflow 的には「保存済みファイルを見る」のか「実行中Excelを見る」のかを分けるべき

仕様としては、backendを明示できるようにすると綺麗です。

```bash
xlflow inspect range --sheet "World News" --address "A1:F20"
```

デフォルトは excelize。

```bash
xlflow inspect range --sheet "World News" --address "A1:F20" --backend file
```

Excel COMで開いているブックを見る。

```bash
xlflow inspect range --sheet "World News" --address "A1:F20" --backend excel
```

または名前をこうしてもよいです。

```bash
--source file
--source excel
```

個人的には `--source` の方がユーザーに伝わりやすいです。

```bash
xlflow inspect range --sheet "World News" --address "A1:F20" --source file
xlflow inspect range --sheet "World News" --address "A1:F20" --source excel
```

## デフォルトは `--source file` が良い

AIエージェント向けには、デフォルトをCOMにしない方が良いです。

理由は単純で、COMは強力ですが不安定要素もあります。

```txt
- Excelプロセスの状態に依存する
- ダイアログやアドインの影響を受ける
- ブックがロックされていると挙動が複雑になる
- 実行環境がWindows + Excelに限定される
- CIやサーバー環境で使いにくい
```

一方、excelizeは保存済みファイルに対して安定して読めます。
AIエージェントが `push -> run -> inspect` する場合も、`run` の最後に保存する設計なら excelize で十分です。

## 重要: `run` 後に保存する仕様なら excelize でかなり戦える

ここが設計上の分岐点です。

`xlflow run` がマクロ実行後にブックを保存するなら、

```bash
xlflow run Main.FetchNews
xlflow inspect used-range --sheet "World News"
```

この `inspect` は excelize で読めます。

つまり、xlflow全体としては、

```txt
VBA実行
↓
Workbook.Save
↓
excelizeでinspect
```

という流れにすれば、COM inspect の必要性はかなり下がります。

AIエージェント向けには、この方が堅いです。

## 仕様案

私ならこうします。

### Phase 1: excelize only

```bash
xlflow inspect sheets --format json

xlflow inspect range \
  --sheet "World News" \
  --address "A1:F20" \
  --format json

xlflow inspect used-range \
  --sheet "World News" \
  --max-rows 100 \
  --max-cols 30 \
  --format markdown
```

内部実装:

```txt
backend: excelize
対象: 保存済み workbook
```

この時点ではCOMなしでよいです。

### Phase 2: sourceオプション追加

```bash
xlflow inspect range \
  --sheet "World News" \
  --address "A1:F20" \
  --source file
```

```bash
xlflow inspect range \
  --sheet "World News" \
  --address "A1:F20" \
  --source excel
```

`--source excel` はCOMで、開いているブックや保存前状態を読む用途です。

### Phase 3: run連携

```bash
xlflow run Main.FetchNews --save
xlflow inspect used-range --sheet "World News"
```

あるいは、`run` 後の保存が既定ならさらに良いです。

```bash
xlflow run Main.FetchNews
xlflow inspect used-range --sheet "World News"
```

## コマンドヘルプには明確に書いた方がいい

ここはユーザーとAIの誤解を防ぐために重要です。

```txt
By default, inspect reads the saved workbook file directly.
Unsaved changes in an open Excel window are not visible.
Use --source excel to inspect the live Excel application state.
```

日本語ならこうです。

```txt
既定では保存済みのブックファイルを直接読み取ります。
Excel上で未保存の変更は反映されません。
開いているExcelの現在状態を読み取りたい場合は --source excel を使用してください。
```

これを明示しないと、「さっきマクロで出力したのにinspectに出ない」という混乱が起きます。

## excelize実装で注意すること

excelizeでやる場合、`UsedRange` 相当は自前定義になります。

Excel COMの `Worksheet.UsedRange` と完全一致させる必要はありません。
むしろ、xlflowでは独自にこう定義した方が良いです。

```txt
inspect used-range は、値・数式・ハイパーリンク・結合セルなどが存在する最小矩形範囲を返す
```

ただしMVPではもっと単純でよいです。

```txt
GetRowsで取得できる行列の範囲を used range とする
```

これで十分です。

ただし、書式だけ付いているセルは拾わない可能性があります。
これはむしろAIエージェント向けには都合が良いです。値のない装飾セルまで拾うとノイズになります。

## 軽量inspectなら、書式だけセルは無視でよい

exstruct的にExcel文書を理解したいなら書式も重要ですが、xlflowのinspectでは不要です。

```txt
xlflow inspect used-range = データ確認用
Excel UsedRange = Excel内部状態
```

として割り切るべきです。

なので名称も、厳密には `used-range` より `data-range` の方が正確かもしれません。

```bash
xlflow inspect data-range --sheet "World News"
```

ただ、Excelユーザーには `used-range` の方が伝わるので、使うなら説明で補足すればよいです。

## 画像・図形・グラフはexcelizeで深追いしない

excelizeでも一部のオブジェクト情報は読めますが、xlflowでそこをやり始めると重くなります。

MVPではやらない方がいいです。

やるとしても、将来的にこれくらいまでです。

```bash
xlflow inspect objects --sheet "World News"
```

```json
{
  "sheet": "World News",
  "pictures": {
    "count": 13
  },
  "charts": {
    "count": 0
  },
  "shapes": {
    "count": 0
  }
}
```

位置・サイズ・中身まで取り始めると exstruct 化します。

## 一番おすすめの設計

```txt
xlflow inspect = file snapshot reader
```

として、最初は excelize で作るのが一番良いです。

仕様はこうです。

```bash
xlflow inspect sheets
xlflow inspect range --sheet <sheet> --address <A1:F20>
xlflow inspect used-range --sheet <sheet>
```

オプション:

```bash
--format json|markdown
--max-rows 100
--max-cols 30
--max-cell-chars 500
--include formulas,hyperlinks
--source file
```

将来用に予約:

```bash
--source excel
```

## 判断基準

| 用途                           | 推奨             |
| ------------------------------ | ---------------- |
| 保存済みブックのセル値を読む   | excelize         |
| シート一覧を読む               | excelize         |
| 範囲の値を読む                 | excelize         |
| 数式文字列を読む               | excelize         |
| ハイパーリンクを読む           | excelize         |
| run後に保存された結果を読む    | excelize         |
| 未保存のExcel状態を読む        | COM              |
| Excelの計算結果を信頼して読む  | COM寄り          |
| 図形・画像・グラフの詳細を読む | COM。ただしMVP外 |
| ActiveSheetやSelectionを見る   | COM              |

## 結論

今のxlflowなら、**まず excelize 実装で良い**です。

理由は、`inspect` の主目的が「AIエージェントがVBA実行後の出力をCLIで確認すること」だからです。
`run` 後に保存する運用にすれば、excelizeだけでかなり実用になります。

ただし、設計上は将来COMに逃がせるように、

```bash
--source file
--source excel
```

または内部的に

```go
type InspectBackend interface {
    ListSheets(...) ...
    ReadRange(...) ...
    ReadUsedRange(...) ...
}
```

のような形にしておくのが安全です。
