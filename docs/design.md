# Excelの「ボタン配置」機能についての設計メモ

xlflow で応えるなら、最初は **「フォームコントロールのボタンをシートに配置して、既存マクロを割り当てる」** ところから始めるのが現実的です。

Excel には大きく分けて次のボタンがあります。

| 種類                         | xlflow対応の現実性 | コメント                                              |
| ---------------------------- | -----------------: | ----------------------------------------------------- |
| フォームコントロールのボタン |               高い | `.OnAction` でマクロ割り当て可能。AI/CLI向き          |
| 図形・Shapeをボタン風に使う  |               高い | 角丸四角形などに `.OnAction` を設定できる             |
| ActiveX CommandButton        |             低〜中 | イベントコード、信頼設定、COM依存が重い。後回し推奨   |
| リボンUI追加                 |             中〜低 | OpenXML/customUI 編集が必要。便利だがスコープが大きい |

最初に狙うべきは、**フォームコントロールボタン**か**Shapeボタン**です。

---

## なぜニーズがあるか

業務VBAでは、ユーザーに「マクロ一覧から実行してください」とは言いにくいです。

実務ではかなりの頻度でこうなります。

```text
このボタンを押したら集計
このボタンを押したらCSV出力
このボタンを押したら帳票作成
このボタンを押したら入力チェック
```

つまり、AIエージェントがVBAを書けるだけでは不十分で、**利用者が実行しやすい入口をシート上に作る**ところまでできると完成度が一段上がります。

xlflow がここまで扱えると、単なる VBA テキスト管理ツールではなく、

> Excel業務アプリをCLI/AIエージェントから組み立てるためのハーネス

に近づきます。

---

## 具体的にできそうな機能

例えば、次のようなコマンドが考えられます。

```bash
xlflow button add \
  --sheet "Menu" \
  --text "集計を実行" \
  --macro "Main.RunAggregation" \
  --cell "B2" \
  --width 160 \
  --height 40
```

または JSON/YAML マニフェストで定義する形です。

```yaml
buttons:
  - sheet: Menu
    text: 集計を実行
    macro: Main.RunAggregation
    cell: B2
    width: 160
    height: 40
    style: primary
```

適用コマンド:

```bash
xlflow apply-ui xlflow.ui.yaml
```

この方向はかなり良いです。

---

## 実装方法の候補

### 1. COM経由でフォームボタンを追加する

Windows + Excel 前提なら、かなり現実的です。

VBA的にはこの系統です。

```vb
Set btn = ws.Buttons.Add(left, top, width, height)
btn.Caption = "集計を実行"
btn.OnAction = "Main.RunAggregation"
```

xlflow が Go 製なら、Excel COM を叩いて同等の処理を行うか、内部的に一時VBA/PowerShell/WSHを使って操作する形になります。

メリット:

- Excel標準の機能なので安定しやすい
- `.OnAction` でマクロ割り当てが簡単
- ユーザー視点で自然
- AIエージェントにも扱いやすい

デメリット:

- Windows + Excel インストール環境前提
- COM操作の待機・例外処理が必要
- 位置指定・重複検出・削除などの設計が必要

---

### 2. Shapeをボタンとして追加する

個人的にはこちらもかなりおすすめです。

Excelでは図形にもマクロを割り当てられます。

```vb
Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, left, top, width, height)
shp.TextFrame2.TextRange.Text = "集計を実行"
shp.OnAction = "Main.RunAggregation"
```

メリット:

- 見た目を整えやすい
- 「メニュー画面」を作りやすい
- `xlflow inspect-gui` や `trace` と相性がよい
- 将来的に ExStruct 的な shape 抽出とも相性がよい

デメリット:

- スタイル指定が増えると設計が膨らむ
- 図形名・重複管理が必要
- フォームボタンより壊れやすいケースはある

---

## 最初に実装するならこの仕様がよさそうです

最小構成はこれで十分です。

```bash
xlflow button add \
  --sheet Menu \
  --cell B2 \
  --text "集計を実行" \
  --macro "Main.RunAggregation"
```

内部では以下を行います。

1. 対象ブックを開く
2. 対象シートを取得、なければエラーまたは `--create-sheet`
3. `cell` の位置から `Left` / `Top` を計算
4. 既存の同名ボタンがあれば更新またはエラー
5. ボタンを配置
6. `OnAction` にマクロ名を設定
7. 保存
8. 配置結果を JSON で返す

出力例:

```json
{
  "ok": true,
  "sheet": "Menu",
  "name": "xlflow_btn_RunAggregation",
  "text": "集計を実行",
  "macro": "Main.RunAggregation",
  "cell": "B2",
  "left": 120,
  "top": 36,
  "width": 160,
  "height": 40
}
```

AIエージェント向けには、この JSON 出力がかなり重要です。

---

## コマンド設計案

個人的には `button` という独立コマンドでもよいですが、将来的には `ui` 名前空間の方が広げやすいです。

```bash
xlflow ui button add
xlflow ui button list
xlflow ui button remove
xlflow ui button update
```

または短くするなら:

```bash
xlflow button add
xlflow button list
xlflow button remove
```

将来的に以下も追加できます。

```bash
xlflow ui menu init
xlflow ui shape add
xlflow ui ribbon add
xlflow ui inspect
```

最初は `button add/list/remove` だけで十分です。

---

## 重要なのは「再実行可能性」です

AIエージェントに使わせるなら、単にボタンを追加できるだけでは足りません。

同じコマンドを何度実行しても破綻しない設計が必要です。

例えば、以下のようにします。

```bash
xlflow button add \
  --id run-aggregation \
  --sheet Menu \
  --cell B2 \
  --text "集計を実行" \
  --macro "Main.RunAggregation"
```

内部の図形名を固定します。

```text
xlflow.button.run-aggregation
```

すでに存在すれば、追加ではなく更新します。

これにより、AIエージェントが何度試行してもボタンが増殖しません。

ここはかなり大事です。

---

## さらに良い仕様: マクロ存在確認

ボタン追加時に、指定マクロが本当に存在するかを検証できると強いです。

```bash
xlflow button add \
  --sheet Menu \
  --cell B2 \
  --text "集計を実行" \
  --macro "Main.RunAggregation" \
  --verify-macro
```

存在しない場合:

```text
error: macro not found: Main.RunAggregation

Candidates:
  - Main.RunAggregationReport
  - Main.Run
  - Report.RunAggregation
```

これは xlflow の `lint` や `inspect` とかなり相性がよいです。

---

## 「ボタンを押す」テストもできると強い

ボタン配置だけでなく、割り当て先マクロの実行確認もあると完成度が高いです。

```bash
xlflow button click --sheet Menu --id run-aggregation
```

ただし、実際にGUI上のボタンをクリックする必要はなく、基本的には `.OnAction` のマクロを解決して `Application.Run` すればよいです。

つまり内部的には:

```text
ボタンを探す
↓
OnAction を読む
↓
Application.Run OnAction
```

これならGUI操作に寄せすぎず、CLIで検証できます。

---

## 優先度としてはかなり高いと思います

今の xlflow の方向性を考えると、これは **実用面でかなり刺さる機能**です。

特に Copilot / Claude Code / Codex に Excel VBA を作らせる場合、最終成果物として、

```text
標準モジュールにマクロがあるだけ
```

よりも、

```text
Menuシートに「実行」ボタンがあり、利用者はそこから操作できる
```

の方が圧倒的に納品物っぽいです。

これはレビュー時の印象も良いです。

---

## ただし ActiveX は最初から狙わない方がいいです

ActiveX の CommandButton は、見た目は業務Excelでよく見ますが、xlflowの初期対応としてはおすすめしません。

理由は以下です。

- イベントプロシージャがシートモジュール側に必要
- セキュリティ設定や信頼済みドキュメントの影響を受けやすい
- 32bit/64bitや環境差の影響を受ける可能性がある
- AIエージェントの自動検証と相性が悪い
- 壊れたときの診断が難しい

最初は **フォームボタン / Shape + OnAction** に限定した方が堅実です。

---

## 最終的なおすすめ

xlflow に入れる価値はかなりあります。

おすすめのロードマップはこうです。

```text
Phase 1:
  xlflow button add/list/remove
  フォームコントロール or Shapeボタン
  OnAction割り当て
  id指定による冪等更新

Phase 2:
  YAML/JSONマニフェスト対応
  xlflow ui apply
  メニューシート生成
  マクロ存在検証

Phase 3:
  button click / button run
  配置済みボタンのOnAction実行
  inspect-gui連携

Phase 4:
  Shapeスタイル
  primary/danger/secondary などの簡易プリセット
  レイアウト自動配置

Phase 5:
  Ribbon UI / ActiveX は必要が見えてから検討
```

個人的には、**`xlflow ui` 系コマンドとして実装する価値はかなり高い**です。

特に `run` / `lint` / `inspect-gui` / `trace` と組み合わせると、

> AIがVBAを書く
> xlflowがブックに反映する
> メニューシートに実行ボタンを作る
> ボタンのOnActionを検証する
> 実行ログを見る

という流れが作れます。

これは Excel VBA のAIエージェント開発体験としてかなり強いです。
