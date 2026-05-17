<p align="center">
    <img width="600" alt="xlflow logo" src="docs/images/logo.png" />
</p>

<p align="center">
  <em>Excel VBA development, rebuilt for CLI-first humans and AI agents.</em>
</p>

<p align="center">
  <a href="README.md">English</a>
  |
  <a href="README.ja.md">日本語</a>
</p>

<div align="center">

![GitHub Release](https://img.shields.io/github/v/release/harumiWeb/xlflow?include_prereleases) ![Scoop](https://img.shields.io/scoop/v/xlflow?bucket=https%3A%2F%2Fgithub.com%2FharumiWeb%2Fscoop-bucket) ![GitHub License](https://img.shields.io/github/license/harumiWeb/xlflow) ![GitHub go.mod Go version](https://img.shields.io/github/go-mod/go-version/harumiWeb/xlflow) [![Ask DeepWiki](https://deepwiki.com/badge.svg)](https://deepwiki.com/harumiWeb/xlflow)

</div>

# :surfing_man: xlflow

**xlflow** は、AIエージェント時代のための Excel VBA 開発フレームワークです。

`.xlsm` ブックに閉じ込められがちな VBA を、ソース管理しやすく、CLI から扱いやすい開発ワークフローに変換します。
VBA のエクスポート、編集、lint、インポート、テスト、trace、実行、差分確認をコマンドラインから行えます。

> [!TIP]
> xlflow は Excel を置き換えるツールではありません。Excel VBA の周囲に CLI ベースの開発ハーネスを用意し、人間・スクリプト・AIエージェントが扱いやすい形にするためのツールです。

## デモ

これらの [サンプル](example) は xlflow を使用して僅かな自然言語指示のみでAIエージェントによって作成されました。

<table>
  <tr>
    <td align="center" width="50%">
      <img src="docs/images/world-news.png" alt="world news" width="100%">
      <sub>NewsAPIを使用して世界のニュースを Excel でまとめるマクロ</sub>
    </td>
    <td align="center" width="50%">
      <img src="docs/images/stock-price.png" alt="stock price" width="100%">
      <sub>株価を取得して Excel に表示するマクロ</sub>
    </td>
  </tr>
  <tr>
    <td align="center" width="50%">
      <img src="docs/images/gen-qrcode.png" alt="generate qrcode" width="100%">
      <sub>セル色表現で QRコードを生成して Excel に表示するマクロ</sub>
    </td>
    <td align="center" width="50%">
      <img src="docs/images/tetris.gif" alt="tetris" width="100%">
      <sub>Excel 上でテトリスをプレイできるマクロ</sub>
    </td>
  </tr>
  <tr>
    <td align="center" width="50%">
      <img src="docs/images/space-invader.gif" alt="space invader" width="100%">
      <sub>ユーザーフォーム上でスペースインベーダーをプレイできるマクロ</sub>
    </td>
    <td align="center" width="50%">
      <img src="docs/images/calendar-picker.png" alt="calendar picker" width="100%">
      <sub>リッチなカレンダーピッカー</sub>
    </td>
  </tr>
</table>

---

## なぜ xlflow が必要か

従来の VBA 開発は、Excel 画面と Visual Basic Editor に強く依存しています。
小さな手作業の修正であれば問題ありませんが、ソース管理、テスト、差分確認、AIエージェントによる修正、再現可能な実行を考えると扱いづらくなります。

| 通常の VBA 開発でつらいこと                    | xlflow でできること                                           |
| ---------------------------------------------- | ------------------------------------------------------------- |
| VBA コードが `.xlsm` の中に閉じ込められている  | `.bas` / `.cls` / `.frm` としてエクスポート・インポートできる |
| マクロの入口が分かりにくい                     | `xlflow macros` で実行可能な `Public Sub` を検出できる        |
| 既存 UserForm の名前が分からない               | `xlflow list forms` で workbook 内 UserForm を列挙できる      |
| 既存 UserForm のレイアウト差分を追いにくい     | `xlflow form snapshot` で Designer state を spec 化できる     |
| 実行エラーの場所や原因が分かりにくい           | 構造化エラー、診断情報、trace log を返せる                    |
| ファイル選択や `MsgBox` が自動実行を止める     | headless 実行前に GUI boundary を検出できる                   |
| workbook の変更をレビューしにくい              | セル値、数式、シート、VBA ソースの差分を確認できる            |
| AIエージェントが Excel UI を安全に操作しにくい | CLI コマンドと安定した JSON 出力を提供できる                  |

```text
pull → edit → push → lint → test/run → trace → diff
```

---

## できること

| 領域               | 機能                                                                                                                            |
| ------------------ | ------------------------------------------------------------------------------------------------------------------------------- |
| ソース管理         | 標準モジュール、クラスモジュール、UserForm、Workbook / Worksheet モジュールをエクスポート・インポート                           |
| 実行               | CLI から型付き引数つきでマクロを実行                                                                                            |
| テスト             | VBA のテスト手続きを検出して実行                                                                                                |
| lint               | `Option Explicit` 不足、`Select` / `Activate`、広すぎるエラー処理、暗黙の Variant、Public module field、対話的処理を検出        |
| GUI 安全性         | ファイル選択、`InputBox`、modal `MsgBox`、UserForm などの自動実行に不向きな境界を検出                                           |
| デバッグ           | trace event と runtime diagnostic を収集                                                                                        |
| 差分確認           | workbook のセル値、数式、シート構成、VBA ソース差分を比較                                                                       |
| AIエージェント連携 | 安定した JSON を返し、Codex / Claude / Cursor / Gemini / GitHub Copilot 風ワークフローなどに使わせるための Skill をインストール |

> [!IMPORTANT]
> xlflow は **Windows-first** のツールです。Workbook 操作には **Microsoft Excel + COM + PowerShell** を使用します。

> [!NOTE]
> Excel COM を使う command は、xlflow 自身が使った bridge host を top-level `bridge` に返します。
> ただし workbook 側の VBA が別途 PowerShell を起動する場合、その host は xlflow の bridge host と一致するとは限りません。`powershell.exe` と `pwsh.exe` の差を追うときは、workbook 側で解決された実行ファイルも確認してください。

---

## 動作要件

| 要件                                                       | 必要になる場面                                                                                                                                                                      |
| ---------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Windows                                                    | Excel COM automation                                                                                                                                                                |
| Microsoft Excel                                            | `new`, `init`, `list forms`, `inspect form`, `form snapshot`, `form build`, `form export-image`, `pull`, `push`, `run`, `export-image`, `edit`, `test`, `macros`, `trace`, `doctor` |
| PowerShell                                                 | Excel automation bridge                                                                                                                                                             |
| VBA プロジェクト オブジェクト モデルへのアクセスを信頼する | VBA プロジェクトの読み書き                                                                                                                                                          |

> [!NOTE]
> `lint`、一部の `diff`、Go のユニットテストなど、Excel COM を使わない処理は非 Excel 環境でも検証できます。

> [!WARNING]
> Excel の設定で **VBA プロジェクト オブジェクト モデルへのアクセスを信頼する** を有効にしてください。これが無効だと、Excel がインストールされていても `pull` / `push` / `run` などが失敗する場合があります。
>
> <details>
> <summary>詳細</summary>
> Excel のオプションで「トラスト センター」→「マクロの設定」→「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を有効にしてください。
> </details>

---

## インストール

### Go install

```bash
go install github.com/harumiWeb/xlflow/cmd/xlflow@latest
```

`go install` は Go 環境に設定された module mirror や checksum database へアクセスすることがあります。source checkout からの開発や CI では、`go.mod` に書かれた Go version を正式サポート toolchain の source of truth としてください。リポジトリの CI / release workflow もその値から Go を解決します。

### Scoop

```powershell
scoop bucket add harumiweb https://github.com/harumiWeb/scoop-bucket
scoop install xlflow
```

### GitHub Releases

Windows 向けの事前ビルド済みバイナリは次のページから取得できます。

[https://github.com/harumiWeb/xlflow/releases](https://github.com/harumiWeb/xlflow/releases)

> [!IMPORTANT]
> 現在の事前ビルド配布は **Windows 向けのみ** です。
> Workbook を操作する command には、**Microsoft Excel**、Excel COM automation、**VBA プロジェクト オブジェクト モデルへのアクセスを信頼する** 設定が必要です。
> Release binary には runtime PowerShell bridge script が埋め込まれているため、`xlflow.exe` 単体で workbook command を実行できます。

ダウンロードした ZIP は、公開されている `checksums.txt` と照合して SHA256 を確認できます。

```powershell
Get-FileHash .\xlflow_windows_x86_64.zip -Algorithm SHA256
certutil -hashfile .\xlflow_windows_x86_64.zip SHA256
```

表示された SHA256 が `checksums.txt` 内の `xlflow_windows_x86_64.zip` の値と一致することを確認してください。

> この確認で分かるのは、ダウンロードしたファイルが公開された checksum file と一致していることです。配布者の本人性を証明するものではなく、Windows の Authenticode signing の代替でもありません。

GitHub CLI では、GitHub Actions provenance attestation も検証できます。

```powershell
gh attestation verify .\xlflow_windows_x86_64.zip --repo harumiWeb/xlflow
```

> この検証で分かるのは、release artifact に対する GitHub artifact attestation が存在し、検証できることです。Windows の publisher certificate による Authenticode signing を意味するものではありません。

インストール後、次のコマンドで確認できます。

```bash
xlflow version
xlflow --help
```

開発中のリポジトリから直接実行する場合:

```bash
go run ./cmd/xlflow --help
```

Taskfile を使用している場合:

```bash
task run -- --help
```

---

## クイックスタート

### 1. プロジェクトを作成または初期化する

新しい xlflow プロジェクトと macro-enabled workbook を作成します。

```bash
xlflow new Book.xlsm
```

既存の Excel ブックから始める場合は `init` を使用します。

```bash
xlflow init Book.xlsm
```

AI エージェント向けの Skill も同時にインストールする場合:

```bash
xlflow new Book.xlsm --with-skill --agent codex
```

interactive な `xlflow new` / `xlflow init` では welcome banner を表示し、最新 GitHub Release を GitHub Releases API で確認することがあります。このリクエストを今回だけ止めたい場合は `--no-update-check`、環境全体で止めたい場合は `XLFLOW_NO_UPDATE_CHECK=1` を使ってください。

### 2. Excel automation 環境を確認する

```bash
xlflow doctor --json
```

> [!TIP]
> `pull` / `push` / `run` / `test` が Excel、COM、PowerShell、VBIDE 設定の問題で失敗する場合は、まず `doctor` を実行してください。

### 3. VBA をソースファイルとして取り出す

```bash
xlflow pull --json
```

エクスポートされた `.bas` / `.cls` / `.frm` は `src/` 配下に出力されます。
folder mode が有効な場合、各 source root 配下のネストしたディレクトリは `push` 時に Rubberduck 互換の `@Folder(...)` annotation へマッピングされます。
通常のエディタや AI エージェントで編集できます。

### 4. 編集したソースを workbook に反映する

```bash
xlflow push --json
```

### 5. マクロを検出して実行する

```bash
xlflow macros --json
xlflow run Main.Run --json
```

無人実行では headless mode を推奨します。

```bash
xlflow run Main.Run --headless --json
```

ファイル選択、MsgBox、UserForm などを人間が操作する場合は interactive mode を使用します。

```bash
xlflow run Main.Run --interactive --timeout 5m --json
```

### 6. lint と test を実行する

```bash
xlflow lint --json
xlflow test --json
```

---

## よく使うワークフロー

### AIエージェントに VBA を編集させる

```text
1. xlflow.toml を読む
2. 通常の編集作業では xlflow session start を実行する
3. どちらが最新か曖昧なら xlflow pull --session --json を実行する
4. src/ 配下の .bas / .cls / .frm を編集する
5. xlflow push --fast --session --no-save --json で workbook へ反映する
6. xlflow lint --json を実行する
7. xlflow test --session --json を実行する
8. xlflow macros --session --json を実行する
9. xlflow run <qualified_name> --headless --session --json を実行する
10. 実行時エラーが分かりにくい場合は xlflow run --trace --session --json を使う
11. xlflow session stop の前に xlflow save --session --json を実行する
12. workbook の変更確認が必要なら xlflow diff --json を使う
```

> [!IMPORTANT]
> AIエージェントや CI 風のスクリプトでは `--json` を使うことを推奨します。JSON envelope は人間向け標準出力よりも安定して扱えるように設計されています。
> `xlflow run` は既定で VBA を compile し、構造化された compile diagnostic を返します。生の Excel / VBE ダイアログを出したい場合だけ `--gui-compile-errors` を使ってください。

### 人間が Excel を操作しながら進める

人間が Excel を開いた状態で作業する場合は、`attach` で active workbook を確認できます。

```bash
xlflow attach --active --json
```

> [!NOTE]
> `attach` は安全確認用です。active workbook が `xlflow.toml` の `excel.path` と一致するかを検証します。`pull` / `push` / `run` の対象を切り替えるコマンドではありません。

### GUI を含むマクロを扱う

headless 実行できるか判断する前に、GUI boundary を確認します。

```bash
xlflow inspect-gui --json
```

| 結果                                                          | 推奨される実行方法                                 |
| ------------------------------------------------------------- | -------------------------------------------------- |
| GUI boundary なし                                             | `xlflow run ... --headless --json`                 |
| ファイル選択、`InputBox`、modal `MsgBox`、UserForm などを検出 | `xlflow run ... --interactive --timeout 5m --json` |
| GUI 処理が実処理を包んでいる                                  | core logic を引数付きの headless 手続きへ分離する  |

> [!WARNING]
> headless automation と modal な Excel UI は相性が悪いです。無人実行前に `inspect-gui` を使い、GUI entrypoint は薄く保つことを推奨します。

---

## コマンドマップ

| コマンド            | 目的                                                    | 代表的な使い方                                                               |
| ------------------- | ------------------------------------------------------- | ---------------------------------------------------------------------------- |
| `new`               | 新しい xlflow プロジェクトと `.xlsm` workbook を作成    | `xlflow new Book.xlsm`                                                       |
| `init`              | 既存 workbook から xlflow プロジェクトを初期化          | `xlflow init Book.xlsm`                                                      |
| `doctor`            | Excel、COM、PowerShell、VBIDE access を診断             | `xlflow doctor --json`                                                       |
| `attach`            | Excel で現在 active な workbook を検証                  | `xlflow attach --active --json`                                              |
| `pull`              | VBA component を `src/` へエクスポート                  | `xlflow pull --json`                                                         |
| `push`              | VBA source を workbook へインポート                     | `xlflow push --json`                                                         |
| `session`           | 高速ループ用に workbook を開いたままにする              | `xlflow session start`                                                       |
| `save`              | session 中の workbook を保存                            | `xlflow save --session --json`                                               |
| `runner`            | 永続 xlflow runner marker module を管理                 | `xlflow runner install --json`                                               |
| `macros`            | 実行可能な macro entrypoint を検出                      | `xlflow macros --json`                                                       |
| `list forms`        | workbook の UserForm と想定 source path を列挙          | `xlflow list forms --json`                                                   |
| `form snapshot`     | 厳密な Designer UserForm state を JSON/YAML spec に保存 | `xlflow form snapshot UserForm1 --out src/forms/specs/UserForm1.yaml --json` |
| `form build`        | 保存済み spec から Designer-backed UserForm を作成      | `xlflow form build src/forms/specs/UserForm1.yaml --json`                    |
| `form export-image` | runtime UserForm を PNG 画像として出力                  | `xlflow form export-image UserForm1 --out artifacts/UserForm1.png --json`    |
| `run`               | CLI から macro を実行                                   | `xlflow run Main.Run --json`                                                 |
| `export-image`      | worksheet range を PNG 画像として出力                   | `xlflow export-image --sheet QR --range A1:AE31 --json`                      |
| `edit`              | live session workbook を準備・調整用に変更する          | `xlflow edit cell --sheet Input --cell B2 --value ABC123 --session --json`   |
| `trace`             | VBA trace log を有効化・収集・削除                      | `xlflow trace enable --json`                                                 |
| `test`              | VBA test を実行                                         | `xlflow test --json`                                                         |
| `diff`              | workbook 内容と任意の VBA source を比較                 | `xlflow diff before.xlsm after.xlsm --json`                                  |
| `inspect`           | Excel COM を使わず保存済み workbook snapshot を確認     | `xlflow inspect range --sheet Result --address A1:F20 --json`                |
| `lint`              | VBA source を lint                                      | `xlflow lint --json`                                                         |
| `analyze`           | Excel を開かず runtime-risk pattern を解析              | `xlflow analyze --json`                                                      |
| `check`             | `lint` / `analyze` / `doctor` をまとめて実行            | `xlflow check --keepalive --json`                                            |
| `inspect-gui`       | GUI interaction boundary を検出                         | `xlflow inspect-gui --json`                                                  |
| `skill install`     | AI エージェント向け Skill をインストール                | `xlflow skill install --agent codex`                                         |
| `version`           | インストール済み xlflow の build metadata を表示        | `xlflow version`                                                             |

---

## コマンド詳細

<details open>
<summary><strong>プロジェクト作成: <code>new</code>, <code>init</code>, <code>doctor</code>, <code>attach</code></strong></summary>

### `xlflow new`

新しい xlflow プロジェクトと `.xlsm` workbook を作成します。

```bash
xlflow new
xlflow new Sales
xlflow new Sales.xlsm
```

引数を省略した場合は `build/Book.xlsm` が作成されます。
拡張子なしの名前を指定した場合は `.xlsm` が付与されます。
`new` は macro-enabled workbook を作成するため、`.xlsm` 以外の拡張子は受け付けません。
scaffolding 時の interactive な GitHub Release 確認を止めたい場合は `--no-update-check` を付けます。

`new` は `xlflow.toml`、`src/`、`tests/`、`build/`、`.xlflow/` などのプロジェクト構造を作成します。
また、Excel 一時ファイルや xlflow の生成物を無視するための `.gitignore` も作成または更新します。

### `xlflow init`

既存の Excel workbook から xlflow プロジェクトを作成します。

```bash
xlflow init Book.xlsm
```

指定した workbook は `build/` 配下へコピーされ、`xlflow.toml` の `[excel].path` に記録されます。
scaffolding 時の interactive な GitHub Release 確認を止めたい場合は `--no-update-check` を付けます。

### `xlflow doctor`

Excel automation の実行環境を診断します。

```bash
xlflow doctor --json
```

Excel がインストールされているか、対象 workbook を開けるか、VBIDE にアクセスできるかを確認します。
ソースファイルが存在する場合は、headless 実行を止める可能性がある GUI boundary も診断情報に含めます。

### `xlflow attach`

現在 Excel で active な workbook を確認します。

```bash
xlflow attach --active --json
```

人間参加型セッションの安全確認に使えます。

</details>

<details open>
<summary><strong>VBA source loop: <code>pull</code>, <code>push</code>, <code>macros</code>, <code>run</code></strong></summary>

### `xlflow pull`

設定された workbook から VBA component をエクスポートします。

```bash
xlflow pull --json
```

標準モジュール、クラスモジュール、UserForm、Workbook / Worksheet などの document module を `src/` 配下へ出力します。`[userform].code_source = "sidecar"` のときだけ、form module に VBA 行がある場合 `src/forms/code/<FormName>.bas` へ code-behind sidecar も出力します。`frm` mode では code-behind は `.frm` に残します。
recorded session workbook を明示的に要求したい場合は `xlflow pull --session --json` を使います。`.xlflow/session.json` が設定済み workbook を指している場合、通常の `xlflow pull --json` でもその live workbook を自動再利用します。
workbook 内に UserForm が検出された場合、`pull` は `.frm` だけでは `.frx` や Designer state を完全には表せないことを warning として返します。

### `xlflow push`

`src/` 配下の VBA source を Excel workbook へインポートします。

```bash
xlflow push --json
```

`.bas` / `.cls` / `.frm` を読み込み、VBIDE 経由で workbook へ反映します。
UserForm の `.frx` は binary companion file として扱います。`[userform].code_source = "sidecar"` のとき、`src/forms/code/*.bas` の UserForm code-behind sidecar は standalone module としては import せず、先に sidecar から tracked `.frm` の埋め込みコードを同期し、その後 `.frm` import 後に対応する UserForm `CodeModule` へ再適用します。`code_source = "frm"` では `.frm` 埋め込みコードが authoritative です。
source 内に UserForm がある場合、`push` はより深い form inspection 用の hint を返します。`push --session --no-save` では live workbook の UserForm state と disk の差分に関する追加 warning も返します。
デフォルトでは `.xlflow/backups` にバックアップを作成し、workbook を保存します。

開発中の反復を速くしたい場合は次を使えます。

```bash
xlflow push --fast --json
xlflow push --changed-only --json
xlflow push --backup=never --json
```

`--fast` は `--backup=never --changed-only` の短縮形です。
`--changed-only` は `.xlflow/state/push.json` の source fingerprint が同じ場合、Excel/VBIDE import をスキップします。

### `xlflow macros`

実行可能な `Public Sub` entrypoint を検出します。

```bash
xlflow macros --json
```

> [!TIP]
> AIエージェントや自動化スクリプトは、macro 名を推測する前にこのコマンドを実行してください。返された `qualified_name` を `xlflow run` に渡すことで、entrypoint の指定ミスを減らせます。
> session 中の開発では、通常の `xlflow macros --json` でも一致する recorded session workbook を自動再利用します。明示的にその workbook を要求したい場合だけ `--session` を付けます。

### `xlflow list forms`

runtime で UserForm を読み込まずに workbook 内の `UserForm` component を列挙します。

```bash
xlflow list forms --json
xlflow list forms --session --json
```

各結果には form 名、想定される `.frx` companion の有無、そして `pull` と同じ folder-aware path ルールで解決した project-relative な `.frm` / `.frx` path が含まれます。

他の workbook-backed な read command と同様に、`.xlflow/session.json` が設定済み workbook を指していれば `list forms` も一致する recorded session workbook を自動再利用します。明示的にその前提を要求したい場合は `--session` を付けます。

### `xlflow inspect form`

Excel COM 経由で workbook の `UserForm` を inspect し、form / control の構造化 state を返します。

```bash
xlflow inspect form UserForm1 --runtime --json
xlflow inspect form UserForm1 --runtime --initializer InitializeForm --json
xlflow inspect form UserForm1 --designer --json
xlflow inspect form UserForm1 --both --initializer InitializeForm --json
```

既定は `--runtime` です。現在の workbook state から作成した一時コピー上で form を読み込み、loaded control state を取得して返却前に unload するため、元の workbook は変更しません。`--designer` は元 workbook の VBIDE Designer state を runtime load なしで読みます。`--both` は両方を返します。

`--initializer <method>` は任意で、`--runtime` または `--both` でのみ使えます。runtime inspect の control 列挙前に、その public form method を `ThisWorkbook` 引数付きで呼びます。`UserForm_Initialize` ではなく独自 initializer で画面 state を埋める form に向いています。

runtime inspect は常に `UserForm_Initialize` 実行 warning を返します。`--initializer` を使った場合は、その explicit initializer を呼んだ warning も返します。

### `xlflow form snapshot`

workbook の `UserForm` について、厳密な design-time snapshot をレビューしやすい JSON / YAML spec ファイルとして保存します。

```bash
xlflow form snapshot UserForm1 --out src/forms/specs/UserForm1.json --json
xlflow form snapshot UserForm1 --out src/forms/specs/UserForm1.yaml --session --json
```

`xlflow inspect form --designer` は source workbook の VBIDE Designer を read-only で直接読み取る経路のまま維持され、workbook VBA の実行を前提としません。一方 `form snapshot` はより厳密です。一時的な workbook copy を開き、注入した VBA helper を実行して、後続の rebuild workflow に必要な具体的な control type まで含む spec を保存します。

`--out` は必須です。出力拡張子と実際のシリアライズ形式は必ず一致し、`.json` は JSON、`.yaml` / `.yml` は YAML を書きます。それ以外の拡張子は Excel を開く前に失敗します。snapshot は helper 実行経路を使うため、workbook の VBA project が helper を実行できない場合は失敗しえます。

persisted artifact の `warnings` は、保存される spec 自体に属する form-local warning だけを対象にします。`save_required` のような operational warning は command envelope と human output に残し、artifact には書き込みません。

継続的に UserForm を管理する場合、Designer 構造の canonical な source-controlled artifact は `src/forms/specs/*.yaml` と考えてください。code-behind の authority は `[userform].code_source` に依存します。新規 project は `sidecar` が既定で `src/forms/code/*.bas` が canonical、`init` した既存 project は `frm` が既定で `.frm` 埋め込みコードが canonical です。`form snapshot` は Designer spec の capture path、`pull` は sidecar mode における code-behind capture path、`form build --overwrite` はその spec から workbook を再構築する rebuild path です。`pull` や `build` で得られる `.frm` / `.frx` は成果物として扱い、Designer behavior の primary source of truth とはみなしません。

他の workbook-backed read command と同様に、`.xlflow/session.json` が設定済み workbook を指していれば `form snapshot` も一致する recorded session workbook を自動再利用します。明示的にその前提を要求したい場合は `--session` を付けます。

### `xlflow form build`

保存済み `xlflow.userform` spec から Designer-backed な workbook `UserForm` を作成します。

```bash
xlflow form build src/forms/specs/UserForm1.yaml --json
xlflow form build src/forms/specs/UserForm1.yaml --session --overwrite --json
```

spec path は `.json` / `.yaml` / `.yml` でなければなりません。xlflow は Excel を開く前に Go 側で schema を検証し、その後 `.frx` を直接編集せず VBIDE Designer API で form と controls を作成します。

spec の parse や validation に失敗した場合、`form build` は `spec_parse_failed`、`spec_validation_failed`、`spec_schema_invalid` のいずれかを返し、top-level `spec` metadata として `path`、`format`、必要に応じて line / column、field path、remediation suggestion を含めます。YAML で `-` や `:` を値として使う場合は quote するか、JSON artifact を使ってください。

既定では同名の `form.name` がすでに存在すると `form_already_exists` で失敗します。`--overwrite` を付けた場合だけ既存 component を削除して spec から作り直します。`src/forms/specs/*.yaml` を source of truth として毎回作り直す置換 workflow はこちらを前提にします。`sidecar` mode では `src/forms/code/<FormName>.bas` から tracked `.frm` artifact を同期し、その sidecar を新しい UserForm へ再適用します。まだ sidecar が無い場合でも削除前 workbook の code-behind を fallback として保持します。`frm` mode では `src/forms/code` を使わず、削除前 workbook の code-behind をそのまま保持します。既定では保存し、`--session --no-save` のときだけ live workbook を未保存のまま残します。この状態では `xlflow save --session` を行うまで live workbook が disk より新しい状態です。

成功した `form build` でも、Designer-backed な弱い項目については contract warning を返すことがあります。form-level の `width` / `height` は best-effort でしかなく、design-time の `ComboBox` / `ListBox` の `list` / `selectedIndex` は xlflow が適用を試みても round-trip 期待としては observed-only とみなしてください。

推奨する UserForm workflow:

```text
1. xlflow list forms --session --json
2. xlflow inspect form <FormName> --designer --session --json
3. xlflow pull --session --json
4. xlflow form snapshot <FormName> --out src/forms/specs/<FormName>.yaml --session --json
5. sidecar mode なら src/forms/code/<FormName>.bas を確認または編集する
6. src/forms/specs/<FormName>.yaml を編集する
7. xlflow form build src/forms/specs/<FormName>.yaml --session --overwrite --json
8. xlflow inspect form <FormName> --designer --session --json
9. xlflow form export-image <FormName> --out artifacts/<FormName>.png --session --json
```

### `xlflow form export-image`

runtime で描画された workbook `UserForm` を PNG 画像として出力します。

```bash
xlflow form export-image UserForm1 --out artifacts/UserForm1.png --json
xlflow form export-image UserForm1 --out artifacts/UserForm1.png --initializer InitializeForm --session --overwrite --json
```

この command は `inspect form --runtime` と同じ安全モデルを使います。現在の workbook state から一時 workbook copy を作成し、その上で form を読み込み、必要なら `ThisWorkbook` を引数に initializer を呼び、modeless 表示した window をキャプチャしてから unload します。capture 自体で source workbook や recorded live session を変更しません。

`--out` は必須で、出力先は `.png` でなければなりません。既存ファイルは `--overwrite` がない限り失敗します。成功時の JSON には top-level `target`、`forms`、`output`、`warnings` が含まれます。`UserForm_Initialize` を実行する warning を常に返し、Windows desktop Excel の GUI capture に依存するため experimental 扱いです。

### `xlflow run`

CLI から macro を実行します。

```bash
xlflow run Main.Run --json
```

引数付き macro も実行できます。

```bash
xlflow run Report.Generate \
  --arg string:fixtures\sample.xlsx \
  --arg int:3 \
  --arg bool:true \
  --json
```

`--arg` は `string:`、`int:`、`bool:` の型付き引数を受け取ります。
空文字は `string:` のみ許可します。

デフォルトでは、`run` は workbook を保存しません。
実行結果を保存する場合は明示的に `--save` または `--save-as` を指定します。

```bash
xlflow run Report.Generate --save --json
xlflow run Report.Generate --save-as build\Result.xlsm --json
```

実行に失敗した場合は、`macro_failed` または `macro_not_found` として、VBA エラー番号、説明、モジュール名、フェーズ、可能であれば行番号を JSON で返します。
近傍ソースや `Set` 抜けなどの既知パターンに一致した場合は、top-level `run_diagnostic` も返します。
既定では `run` が実行前に VBA project を compile し、取得できた module、行、列、message、近傍コードを `vba_compile_failed` として構造化して返します。生の Excel / VBE compile ダイアログをそのまま出したい場合だけ `--gui-compile-errors` を使います。

| Mode                   | 挙動                                                                                                                |
| ---------------------- | ------------------------------------------------------------------------------------------------------------------- |
| `--headless`           | Excel 起動前に GUI boundary を検出し、見つかった場合は `gui_boundary_detected` と top-level `gui_boundaries` を返す |
| `--interactive`        | Excel を表示し、alerts を有効にして人間が操作できる状態で実行する                                                   |
| `--direct`             | 引数なし・trace 無効の macro を temporary harness 注入なしで実行する。単独指定なら既定diagnosticを自動で無効化する  |
| `--fast`               | diagnostic を無効化したうえで条件を満たす場合は direct 実行し、それ以外は通常実行へ戻る                             |
| `--diagnostic`         | 構造化 compile diagnostic を有効のまま明示する（既定で true）                                                       |
| `--gui-compile-errors` | 構造化 compile diagnostic を opt-out し、Excel / VBE の compile dialog を表示する                                   |
| `--session`            | `xlflow session start` で開いた workbook を使う                                                                     |
| `--timeout 5m`         | 指定時間内に終わらない場合は停止し、`macro_timeout` を返す                                                          |

### `xlflow session`

Excel と設定済み workbook を開いたままにします。

```bash
xlflow session start
xlflow pull --session --json   # workbook 側が新しいかもしれない場合
xlflow push --fast --session --no-save --json
xlflow run Main.Run --headless --session --json
xlflow save --session --json
xlflow session stop
```

`--session` は明示的な強制 attach 用に残ります。`.xlflow/session.json` が設定済み workbook を指している場合は、通常の `list forms` / `inspect form` / `form snapshot` / `form build` / `pull` / `push` / `macros` / `run` / `export-image` / `form export-image` / `test` / `trace` / `save` でも一致する live workbook を自動再利用し、その利用形態を JSON / human output に表示します。

`push --session --no-save` が成功した場合や、`run --session` を `--save` / `--save-as` なしで実行した場合は、`xlflow save --session` を行うまで live workbook とディスク上の `.xlsm` がずれる可能性があります。
UserForm を含む workbook では、`.frm` / `.frx` 差分を確認する前にその save を行う前提で考えてください。
xlflow はこの未保存 session 状態を以前より強く警告しますが、`session stop` 前に明示的に `xlflow save --session` する運用が基本です。

</details>

<details open>
<summary><strong>デバッグとテスト: <code>trace</code>, <code>test</code>, <code>diff</code></strong></summary>

### `xlflow trace`

実行中の VBA から log event を収集する仕組みです。

trace module を workbook と source tree に永続化したい場合は有効化します。

```bash
xlflow trace enable --json
xlflow trace status --json
```

session 中の開発では、`trace enable`、`trace status`、`trace disable`、trace 付き `run` に `--session` を付けます。

VBA 側では次のようにログを出せます。

```vb
Call XlflowLog("start GenerateReport")
Call XlflowLog("lastRow=" & lastRow)
Call XlflowLog("finished GenerateReport")
```

trace を有効にして macro を実行します。

```bash
xlflow run Main.Run --trace --json
```

trace event は JSON の top-level `trace` フィールドに返されます。
実行時エラーの直前にどこまで処理が進んだかを把握しやすくなります。

`xlflow run --trace` は helper がない場合でも一時的に注入し、実行後に戻します。
human output と JSON の `trace.lifecycle` で、その一時注入か、すでに永続化済みの helper かを区別できます。
trace log は `.xlflow/traces` に保存されます。
永続化した helper を外すには `xlflow trace disable --json`、ログを消すには `xlflow trace clean --json` を使います。
`xlflow trace inject` は互換用 alias として残ります。

### `xlflow test`

VBA test を実行します。

```bash
xlflow test --json
```

`Test` で始まる、または `_Test` で終わる引数なしの `Sub` を test として検出します。
xlflow session が開いている場合は `xlflow test --session --json` を使います。

特定の test だけを実行する場合は `--filter` を指定します。

```bash
xlflow test --filter TestCreateReport --json
```

新規作成または初期化されたプロジェクトには `src/modules/XlflowAssert.bas` が含まれます。
`AssertEquals expected, actual, [message]` を使って scalar value を比較できます。

```vb
Public Sub TestCreateReport()
    AssertEquals 10, Sheets("Result").Range("A1").Value2
End Sub
```

> [!NOTE]
> `AssertEquals` は object や array の比較には対応していません。`Range` object そのものではなく、`Range.Value2` などの scalar property を比較してください。

### `xlflow diff`

2つの workbook を比較します。

```bash
xlflow diff before.xlsm after.xlsm --json
```

シートの追加・削除、セル値、数式の差分を検出します。

VBA source の差分も比較する場合は、`--vba-before` と `--vba-after` を指定します。

```bash
xlflow diff before.xlsm after.xlsm \
  --vba-before before-src \
  --vba-after after-src \
  --json
```

> [!IMPORTANT]
> 差分が見つかった場合でも、コマンド自体は成功として扱われます。`diff` は差分があっても exit code `0` を返します。差分の有無は JSON の `diff.summary.total_diffs` を確認してください。

</details>

<details open>
<summary><strong>品質ゲート: <code>lint</code>, <code>analyze</code>, <code>check</code>, <code>inspect-gui</code></strong></summary>

### `xlflow lint`

VBA source を lint します。

```bash
xlflow lint --json
```

AIエージェントや無人実行と相性の悪い書き方を検出します。

- `Option Explicit` の不足
- `Select` の使用
- `Activate` の使用
- `On Error Resume Next` の使用
- 暗黙の `Variant` の疑い
- module-level `Public` 変数
- `Application.GetOpenFilename`、`Application.FileDialog`、`InputBox`、modal `MsgBox` などの対話的処理

### `xlflow analyze`

Excel を開かずに、VBA の実行時エラーにつながりやすい pattern を解析します。

```bash
xlflow analyze --json
```

最初の analyzer rule は、object 変数や object を返す function に対する `Set` 抜けを検出します。
結果は top-level `analysis` に、file、module、procedure、line、nearby code、reason、suggestion として返ります。

### `xlflow check`

標準の事前確認をまとめて実行します。

```bash
xlflow check --keepalive --json
```

`check` は `lint`、`analyze`、`doctor` を順に実行し、top-level `check` に集約結果を返します。
lint/analyze の finding があっても続行するため、Excel COM の doctor 結果まで含めて確認できます。

### `xlflow inspect`

Excel を開かず、保存済みの workbook ファイルを直接読み取ります。

```bash
xlflow inspect workbook --json
xlflow inspect sheets --format markdown
xlflow inspect range --sheet "Result" --address "A1:F20" --json
xlflow inspect range --sheet "QR" --address "A1:AE31" --include-style --json
xlflow inspect used-range "Result" --max-rows 50 --max-cols 10 --format markdown
xlflow inspect cell "Result!B3" --json
```

`push` / `run` 後に保存済み workbook の構造やセル出力を確認したいときに使います。
`src/forms` 配下に `.frm` がある場合、`inspect` は live UserForm の Designer/runtime state を確認していないことを warning で明示します。
`inspect` は file snapshot reader なので、Excel 上で未保存の変更はこのコマンド群では見えません。
塗りつぶし色、罫線、結合セル、行高、列幅まで確認したい場合は、`inspect range` または `inspect used-range` に `--include-style` を付けます。

### `xlflow export-image`

Excel COM 経由で worksheet range を PNG として出力します。

```bash
xlflow export-image --sheet "QR" --range "A1:AE31" --json
xlflow export-image --sheet "QR" --range "A1:AE31" --out artifacts\qr.png --overwrite --json
```

`inspect` の補完として使います。チャート、塗りつぶし、レイアウト、帳票、QRコードのセル描画など、保存済み workbook snapshot だけでは十分に検証できない見た目を確認したいときに向いています。

`--out` を省略した場合は `.xlflow/artifacts/images/<workbook-name>/` 配下へ生成名で保存します。`--output-dir` はディレクトリだけ、`--name` はファイル名だけを指定します。v1 では PNG のみ対応です。

他の workbook-backed command と同様に、`.xlflow/session.json` が設定済み workbook を指していれば `export-image` も一致する recorded session workbook を自動再利用します。明示的にその前提を要求したい場合は `--session` を付けます。成功時の JSON には top-level `target`、`output`、必要に応じて `warnings` が含まれます。

### `xlflow edit`

xlflow session で開いている live workbook を、開発中の入力準備や見た目調整のために直接変更します。

```bash
xlflow edit cell --sheet "Input" --cell "B2" --value "ABC123" --events on --session --json
xlflow edit range --sheet "QR" --range "A1:AE31" --fill "#FFFFFF" --session --json
xlflow edit rows --sheet "QR" --rows "1:31" --height 12 --session --json
xlflow edit columns --sheet "QR" --columns "A:AE" --width 2.2 --session --json
```

`run` 前の workbook state 準備、`--events on` による `Worksheet_Change` の再現、再試行前の範囲クリア、画像出力前の行高・列幅調整に使えます。MVP の `edit` は意図的に session-only で、`--session` が必須です。成功した edit は live workbook を dirty にするため、disk へ残すには `xlflow save --session` を実行してください。

### `xlflow inspect-gui`

Excel を開かずに GUI boundary を検出します。

```bash
xlflow inspect-gui --json
```

結果には file、line、kind、symbol、suggestion が含まれます。
macro を `--headless` で実行できるか、`--interactive` が必要かを判断する材料になります。

</details>

<details open>
<summary><strong>AIエージェント対応: <code>skill install</code></strong></summary>

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

| Agent target | Install path            |
| ------------ | ----------------------- |
| `agents`     | `.agents/skills/xlflow` |
| `codex`      | `.codex/skills/xlflow`  |
| `claude`     | `.claude/skills/xlflow` |
| `cursor`     | `.cursor/skills/xlflow` |
| `gemini`     | `.gemini/skills/xlflow` |

GitHub Copilot 風の workflow で使わせる場合は、共通の `.agents` target を使用してください。

```bash
xlflow skill install --agent agents
```

</details>

---

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

[vba]
folders = true
folder_annotation = "update"
default_component_folders = true

[userform]
code_source = "sidecar"

[lint]
require_option_explicit = true
forbid_select = true
forbid_activate = true
forbid_on_error_resume_next = true
detect_implicit_variant = true
forbid_public_module_fields = true
forbid_interactive_input = true
```

`project.entry` は `xlflow run` の macro 名を省略した場合に使われます。

対話前提の project で `UserForm` やダイアログを意図的に使う場合は、`forbid_interactive_input = false` にすると `VB007` 警告を抑止できます。これは lint だけに効き、`xlflow run --headless` の GUI 境界チェックは引き続きブロックします。

typographic quote、C-style quote escape、閉じられていないまたは対応がずれた procedure、行継続 `_` の空白不足を検出する構文安全 lint は常に有効です。`push` や `run` が Excel を開く前に VBE compile dialog を防ぐためのルールです。

---

## JSON 出力

すべてのコマンドは `--json` を付けることで、AIエージェントやスクリプトから扱いやすい JSON を返します。

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

> [!TIP]
> AIエージェントや自動化スクリプトでは、`status`、`command`、`error.code`、各コマンド固有の top-level field を主な contract として扱うことを推奨します。

---

## Exit code

| Code | 意味                                           |
| ---: | ---------------------------------------------- |
|  `0` | 成功                                           |
|  `1` | lint、macro、test などの検証失敗               |
|  `2` | CLI 引数または設定エラー                       |
|  `3` | Excel、COM、VBIDE、PowerShell などの環境エラー |

> [!NOTE]
> `diff` は差分が見つかった場合でも exit code `0` を返します。差分の有無は `diff.summary.total_diffs` を確認してください。

---

## VBAを書くときの推奨ルール

xlflow で実行する VBA は、無人実行しやすい形に寄せることを推奨します。

- [x] 必ず `Option Explicit` を使う
- [x] `Workbook` / `Worksheet` / `Range` を明示的に参照する
- [x] `Integer` より `Long` を使う
- [x] GUI entrypoint は薄く保つ
- [x] 実処理は引数付きの headless 実行可能な手続きへ分離する
- [x] 入力値は `xlflow run --arg`、設定ファイル、固定パス、環境変数などから渡す
- [x] エラーハンドラで原因が分かるメッセージを出す
- [x] workbook を破壊的に変更する処理は test または diff で検証する
- [ ] `Select` / `Activate` / `ActiveSheet` に依存しない
- [ ] headless 手続きで UI ダイアログや modal `MsgBox` に依存しない
- [ ] `On Error Resume Next` を広範囲に使わない

---

## ローカル検証

リポジトリの lint は次のコマンドで実行できます。

```bash
task lint
```

`task lint` は `golangci-lint run` と、追跡対象の `.ps1` source に対する `PSScriptAnalyzer` を実行します。
ローカルの PowerShell 環境で `Invoke-ScriptAnalyzer` が使える状態にしてください。

リポジトリの高速検証は次のコマンドで実行できます。

```bash
task verify
```

現在の `task verify` は、Excel COM に依存しない test として `go test ./...` を実行します。

Excel COM を含む E2E 検証は、Windows + Excel + VBIDE access enabled の環境で実行してください。

```bash
xlflow doctor --json
```

`doctor` が正常であることを確認してから、実 workbook に対して `new` / `doctor` / `pull` / `lint` / `push` / `run` / `test` / `diff` を実行します。

---

## 現在の位置づけ

xlflow は MVP 段階のツールです。

主な目的は、Excel VBA を AIエージェントや CLI ベースの開発フローに載せることです。
特に、次のような用途を想定しています。

| 用途                          | xlflow が役立つ理由                                         |
| ----------------------------- | ----------------------------------------------------------- |
| 既存 VBA のソース管理         | VBA module を通常の file として扱える                       |
| AIエージェントによる VBA 修正 | Agent が source を編集し、check を実行し、JSON 出力を読める |
| Excel macro の CLI 実行       | Terminal や script から macro を起動できる                  |
| VBA の自動テスト              | Test を検出して一貫した形で実行できる                       |
| 実行ログによるデバッグ        | Trace event でどこまで進んだかを確認できる                  |
| Workbook 変更の差分確認       | `diff` で workbook の変更をレビューしやすくなる             |
| 社内 Excel 自動化の保守性向上 | 既存 VBA 資産をより安全な開発 workflow に寄せられる         |

> [!CAUTION]
> xlflow は便利ですが、すべての legacy workbook を自動的に安全な headless 実行へ変換できるわけではありません。GUI-heavy な macro、workbook-level side effect、外部依存、壊れやすい Excel state には意図的なリファクタリングが必要です。

---

## License

MIT
