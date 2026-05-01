<p align="center">
    <img width="600" alt="logo" src="docs/images/logo.png" />
</p>

<p align="center">
  <em>xlflow - An Excel VBA development framework for the AI agent era</em>
</p>

<p align="center">
  <a href="README.md">
    English
  </a>
   |
  <a href="README.ja.md">
    日本語
  </a>
</p>

# xlflow

**AIエージェント時代のための Excel VBA 開発フレームワーク**

xlflow は、Excel VBA プロジェクトを CLI 中心の開発ワークフローに変換するためのツールです。

従来の Excel VBA 開発では、VBE 上で直接コードを書き、手作業でマクロを実行し、問題が起きたら Excel 画面上で原因を探す必要がありました。これは人間にとっても扱いづらく、CLI で開発を進める AI エージェントにとっては特に不向きです。

xlflow は、VBA を通常のソースコードとして管理し、CLI から検査・反映・実行・テスト・差分確認できるようにします。

## できること

- `.xlsm` から VBA モジュールをエクスポートする
- `.bas` / `.cls` / `.frm` を通常のソースコードとして編集する
- 編集した VBA ソースを Excel ブックへインポートする
- CLI からマクロを実行する
- CLI から VBA テストを実行する
- Workbook のセル値・数式・VBA ソース差分を比較する
- VBA を lint して、自動実行に不向きな書き方を検出する
- GUI 操作が必要な境界を検出する
- 実行ログを trace として収集する
- AI エージェント向けに安定した JSON を返す
- Codex / Claude / Cursor / Gemini などに使わせるための Skill をインストールする

## なぜ xlflow が必要か

Excel VBA は、業務現場では今でも重要な自動化基盤です。一方で、AIエージェントにとっては扱いにくい対象でもあります。

主な理由は以下です。

- VBA のコードが `.xlsm` の中に閉じ込められている
- CLI から編集・実行・検証しにくい
- マクロの入口が分かりにくい
- 実行エラーの場所や原因が分かりにくい
- ファイル選択ダイアログや `MsgBox` などが自動実行を止める
- 差分確認やテストの仕組みを作りにくい

xlflow は、これらの問題を解決するために、Excel VBA に対して次のような開発ループを提供します。

```text
pull → edit → push → lint → test/run → trace → diff
```

これにより、人間も AI エージェントも、Excel VBA を通常のソフトウェア開発に近い形で扱えるようになります。

## 動作要件

xlflow は Windows-first のツールです。

Excel 操作には PowerShell と Excel COM を使用します。`pull` / `push` / `run` / `test` / `macros` / `trace` など、Excel ブックを操作するコマンドを実行するには、Windows 環境と Microsoft Excel が必要です。

また、VBA プロジェクトの読み書きには Excel の設定で **VBA プロジェクト オブジェクト モデルへのアクセスを信頼する** を有効にする必要があります。

一方で、`lint` や一部の `diff`、Go のユニットテストなど、Excel COM を使わない処理は非 Excel 環境でも検証できます。

## インストール

Go が利用できる環境では、次のコマンドでインストールできます。

```bash
go install github.com/harumiWeb/xlflow/cmd/xlflow@latest
```

インストール後、次のコマンドで動作確認します。

```bash
xlflow --help
```

開発中のリポジトリから直接実行する場合は、次のように実行できます。

```bash
go run ./cmd/xlflow --help
```

Taskfile を使用している場合は、次のコマンドも利用できます。

```bash
task run -- --help
```

## クイックスタート

新しい xlflow プロジェクトと macro-enabled workbook を作成します。

```bash
xlflow new Book.xlsm
```

AI エージェント向けの Skill も同時にインストールする場合は、次のようにします。

```bash
xlflow new Book.xlsm --with-skill --agent codex
```

既存の Excel ブックから始める場合は、`init` を使用します。

```bash
xlflow init Book.xlsm
```

Excel / COM / VBIDE の状態を確認します。

```bash
xlflow doctor --json
```

ブック内の VBA をソースファイルとしてエクスポートします。

```bash
xlflow pull --json
```

VBA ソースを編集したあと、ブックへ反映します。

```bash
xlflow push --json
```

実行可能なマクロの入口を確認します。

```bash
xlflow macros --json
```

マクロを実行します。

```bash
xlflow run Main.Run --json
```

無人実行では headless mode を推奨します。

```bash
xlflow run Main.Run --headless --json
```

ファイル選択、MsgBox、UserForm などを人間が操作する場合は interactive mode を使います。

```bash
xlflow run Main.Run --interactive --timeout 5m --json
```

VBA テストを実行します。

```bash
xlflow test --json
```

lint を実行します。

```bash
xlflow lint --json
```

## 基本コマンド

### `xlflow new`

新しい xlflow プロジェクトと `.xlsm` ブックを作成します。

```bash
xlflow new
xlflow new Sales
xlflow new Sales.xlsm
```

引数を省略した場合は `build/Book.xlsm` が作成されます。拡張子なしの名前を指定した場合は `.xlsm` が付与されます。`new` は macro-enabled workbook を作成するため、`.xlsm` 以外の拡張子は受け付けません。

`new` は `xlflow.toml`、`src/`、`tests/`、`build/`、`.xlflow/` などのプロジェクト構造を作成します。また、Excel 一時ファイルや xlflow の生成物を無視するための `.gitignore` も作成または更新します。

### `xlflow init`

既存の Excel ブックから xlflow プロジェクトを作成します。

```bash
xlflow init Book.xlsm
```

指定したブックは `build/` 配下へコピーされ、`xlflow.toml` の `[excel].path` に記録されます。

### `xlflow doctor`

Excel automation の実行環境を診断します。

```bash
xlflow doctor --json
```

Excel がインストールされているか、対象ブックを開けるか、VBIDE にアクセスできるかを確認します。`pull` / `push` / `run` / `test` が環境起因で失敗する場合は、まず `doctor` を実行してください。

ソースファイルが存在する場合は、headless 実行を止める可能性がある GUI boundary も診断情報に含めます。

### `xlflow attach`

現在 Excel で active な workbook を確認します。

```bash
xlflow attach --active --json
```

人間参加型セッションの安全確認用です。active workbook が `xlflow.toml` の `excel.path` と一致するかを検証します。`pull` / `push` / `run` の対象を切り替える機能ではありません。

### `xlflow pull`

設定されたブックから VBA コンポーネントをエクスポートします。

```bash
xlflow pull --json
```

標準モジュール、クラスモジュール、UserForm、Workbook / Worksheet などのドキュメントモジュールを `src/` 配下へ出力します。

### `xlflow push`

`src/` 配下の VBA ソースを Excel ブックへインポートします。

```bash
xlflow push --json
```

`.bas` / `.cls` / `.frm` を読み込み、VBIDE 経由でブックへ反映します。UserForm の `.frx` はバイナリ companion file として扱います。

### `xlflow macros`

実行可能な Public Sub を検出します。

```bash
xlflow macros --json
```

AI エージェントや自動化スクリプトは、マクロ名を推測する前にこのコマンドを実行してください。`qualified_name` に表示された名前を `xlflow run` に渡すことで、入口名の誤りを減らせます。

### `xlflow run`

CLI からマクロを実行します。

```bash
xlflow run Main.Run --json
```

引数付きマクロも実行できます。

```bash
xlflow run Report.Generate \
  --arg string:fixtures\sample.xlsx \
  --arg int:3 \
  --arg bool:true \
  --json
```

`--arg` は `string:`、`int:`、`bool:` の型付き引数を受け取ります。`string:` のみ空文字を許可します。

デフォルトでは、`run` はブックを保存しません。実行結果を保存する場合は明示的に `--save` または `--save-as` を指定します。

```bash
xlflow run Report.Generate --save --json
xlflow run Report.Generate --save-as build\Result.xlsm --json
```

実行に失敗した場合は、`macro_failed` または `macro_not_found` として、VBA エラー番号、説明、モジュール名、フェーズ、可能であれば行番号を JSON で返します。近傍ソースや `Set` 抜けなどの既知パターンに一致した場合は、top-level `run_diagnostic` も返します。

`--headless` は Excel 起動前に GUI boundary を検出し、見つかった場合は `gui_boundary_detected` と top-level `gui_boundaries` を返して失敗します。`--interactive` は Excel を表示し、alerts を有効にして人間が操作できる状態で実行します。`--timeout` は既定で `5m` です。

### `xlflow trace`

実行中の VBA からログイベントを収集するための仕組みです。

trace helper を永続化したい場合は有効化します。

```bash
xlflow trace enable --json
xlflow trace status --json
```

VBA 側では次のようにログを出せます。

```vb
Call XlflowLog("start GenerateReport")
Call XlflowLog("lastRow=" & lastRow)
Call XlflowLog("finished GenerateReport")
```

trace を有効にしてマクロを実行します。

```bash
xlflow run Main.Run --trace --json
```

trace event は JSON の top-level `trace` フィールドに返されます。実行時エラーの直前にどこまで処理が進んだかを把握しやすくなります。

`xlflow run --trace` は helper がない場合でも一時的に注入し、実行後に戻します。trace log は `.xlflow/traces` に保存されます。永続化した helper を外すには `xlflow trace disable --json`、ログを消すには `xlflow trace clean --json` を使います。`xlflow trace inject` は互換用 alias として残ります。

### `xlflow test`

VBA テストを実行します。

```bash
xlflow test --json
```

`Test` で始まる、または `_Test` で終わる引数なしの `Sub` をテストとして検出します。

特定のテストだけを実行する場合は `--filter` を指定します。

```bash
xlflow test --filter TestCreateReport --json
```

新規作成または初期化されたプロジェクトには `src/modules/XlflowAssert.bas` が含まれます。`AssertEquals expected, actual, [message]` を使って scalar value を比較できます。

```vb
Public Sub TestCreateReport()
    AssertEquals 10, Sheets("Result").Range("A1").Value2
End Sub
```

`AssertEquals` はオブジェクトや配列の比較には対応していません。`Range` オブジェクトそのものではなく、`Range.Value2` などの scalar property を比較してください。

### `xlflow diff`

2つの Workbook を比較します。

```bash
xlflow diff before.xlsm after.xlsm --json
```

シートの追加・削除、セル値、数式の差分を検出します。

VBA ソースの差分も比較する場合は、`--vba-before` と `--vba-after` を指定します。

```bash
xlflow diff before.xlsm after.xlsm \
  --vba-before before-src \
  --vba-after after-src \
  --json
```

差分が見つかった場合でも、コマンド自体は成功として扱われます。差分の有無は JSON の `diff.summary.total_diffs` を確認してください。

### `xlflow lint`

VBA ソースを lint します。

```bash
xlflow lint --json
```

次のような、AIエージェントや自動実行と相性の悪い書き方を検出します。

- `Option Explicit` の不足
- `Select` の使用
- `Activate` の使用
- `On Error Resume Next` の使用
- 暗黙の `Variant` の疑い
- module-level `Public` 変数
- `Application.GetOpenFilename`、`Application.FileDialog`、`InputBox`、modal `MsgBox` などの対話的処理

### `xlflow analyze`

Excel を開かずに、VBA の実行時エラーにつながりやすいパターンを解析します。

```bash
xlflow analyze --json
```

最初の analyzer rule は、object 変数や object を返す function に対する `Set` 抜けを検出します。結果は top-level `analysis` に、file、module、procedure、line、nearby code、reason、suggestion として返ります。

### `xlflow check`

標準の事前確認をまとめて実行します。

```bash
xlflow check --keepalive --json
```

`check` は `lint`、`analyze`、`doctor` を順に実行し、top-level `check` に集約結果を返します。lint/analyze の finding があっても続行するため、Excel COM の doctor 結果まで含めて確認できます。

### `xlflow inspect-gui`

Excel を開かずに GUI boundary を検出します。

```bash
xlflow inspect-gui --json
```

結果には file、line、kind、symbol、suggestion が含まれます。マクロを `--headless` で実行できるか、`--interactive` が必要かを判断する材料になります。

### `xlflow skill install`

AI エージェント向けの xlflow Skill をインストールします。

```bash
xlflow skill install --agent codex
xlflow skill install --agent claude
xlflow skill install --agent cursor
xlflow skill install --agent gemini
xlflow skill install --target .agents/skills
```

対応している provider target は次の通りです。

- `agents`: `.agents/skills/xlflow`
- `codex`: `.codex/skills/xlflow`
- `claude`: `.claude/skills/xlflow`
- `cursor`: `.cursor/skills/xlflow`
- `gemini`: `.gemini/skills/xlflow`

GitHub Copilot に使わせる場合は、共通の `.agents` target を使用してください。

```bash
xlflow skill install --agent agents
```

## 設定ファイル

xlflow はプロジェクトルートの `xlflow.toml` を読み込みます。

```toml
[project]
name = "sample"
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"
visible = false
display_alerts = false

[src]
modules = "src/modules"
classes = "src/classes"
forms = "src/forms"
workbook = "src/workbook"

[lint]
require_option_explicit = true
forbid_select = true
forbid_activate = true
forbid_on_error_resume_next = true
detect_implicit_variant = true
forbid_public_module_fields = true
forbid_interactive_input = true
```

`project.entry` は `xlflow run` のマクロ名を省略した場合に使われます。

## JSON 出力

すべてのコマンドは `--json` を付けることで、AI エージェントやスクリプトから扱いやすい JSON を返します。

基本的な envelope は次の形式です。

```json
{
  "status": "ok",
  "command": "lint",
  "error": null,
  "logs": []
}
```

失敗時は `status` が `failed` になり、`error.code` と `error.message` が返ります。

```json
{
  "status": "failed",
  "command": "run",
  "error": {
    "code": "macro_failed",
    "message": "Main Err 5: inputPath is required",
    "source": "Main",
    "number": 5,
    "phase": "invoke_macro"
  },
  "logs": []
}
```

AI エージェントや CI は、人間向けの標準出力を parse せず、必ず `--json` を使用してください。

## Exit code

xlflow の exit code は次のように分類されます。

| Code | 意味                                           |
| ---: | ---------------------------------------------- |
|    0 | 成功                                           |
|    1 | lint、macro、test などの検証失敗               |
|    2 | CLI 引数または設定エラー                       |
|    3 | Excel、COM、VBIDE、PowerShell などの環境エラー |

`diff` は差分が見つかった場合でも exit code `0` を返します。差分の有無は `diff.summary.total_diffs` を確認してください。

## AI エージェントに使わせる場合

xlflow は AI エージェントが Excel VBA を安全に編集・実行・検証するための proof loop を提供します。

推奨ワークフローは次の通りです。

```text
1. xlflow.toml を読む
2. 必要なら xlflow pull --json で現在の VBA を取得する
3. src/ 配下の .bas / .cls / .frm を編集する
4. xlflow push --json でブックへ反映する
5. xlflow lint --json で危険な書き方を直す
6. xlflow test --json でテストする
7. テストがない場合は xlflow macros --json → xlflow run <qualified_name> --headless --json を実行する
8. 実行時エラーが分かりにくい場合は xlflow run --trace --json を使い、run_diagnostic と trace output を確認する
9. workbook の変更確認が必要なら xlflow diff --json を使う
```

AI エージェントに渡す場合は、`xlflow skill install` または `xlflow new/init --with-skill` を使って bundled Skill をインストールしてください。

```bash
xlflow skill install --agent codex
```

## VBAを書くときの推奨ルール

xlflow で実行する VBA は、無人実行しやすい形に寄せることを推奨します。

- 必ず `Option Explicit` を使う
- `Select` / `Activate` / `ActiveSheet` に依存しない
- `Workbook` / `Worksheet` / `Range` を明示的に参照する
- `Integer` より `Long` を使う
- UI ダイアログや modal `MsgBox` に依存しない
- GUI 入口は薄く保ち、実処理は引数付きの headless 実行可能な手続きへ分離する
- 入力値は `xlflow run --arg`、設定ファイル、固定パス、環境変数などから渡す
- `On Error Resume Next` を広範囲に使わない
- エラーハンドラで原因が分かるメッセージを出す
- workbook を破壊的に変更する処理はテストまたは diff で検証する

## ローカル検証

リポジトリの高速検証は次のコマンドで実行できます。

```bash
task verify
```

現在の `task verify` は、Excel COM に依存しないテストとして `go test ./...` を実行します。

Excel COM を含む E2E 検証は、Windows + Excel + VBIDE access enabled の環境で実行してください。

```bash
xlflow doctor --json
```

`doctor` が正常であることを確認してから、実ブックに対して `new` / `doctor` / `pull` / `lint` / `push` / `run` / `test` / `diff` を実行します。

## 現在の位置づけ

xlflow は MVP 段階のツールです。

主な目的は、Excel VBA を AI エージェントや CLI ベースの開発フローに載せることです。特に、次のような用途を想定しています。

- 既存 VBA のソース管理
- AI エージェントによる VBA 修正
- Excel マクロの CLI 実行
- VBA の自動テスト
- 実行ログによるデバッグ
- Workbook 変更の差分確認
- 社内 Excel 自動化の保守性向上

## License

MIT
