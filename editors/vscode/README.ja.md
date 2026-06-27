# xlflow for Visual Studio Code

[English](README.md) | [日本語](https://github.com/harumiWeb/xlflow/blob/main/editors/vscode/README.ja.md)

**xlflow for Visual Studio Code**は、Excel VBAマクロ開発支援ツール [xlflow](https://github.com/harumiWeb/xlflow) を VSCode から扱いやすくするための拡張機能です。

xlflowプロジェクトの状態確認、VBAモジュールの取り込み・反映、セッション管理、各種コマンド実行を VSCode 上から行えるようにし、Excel VBAマクロをより安全に、Git管理しやすく、AIエージェントとも連携しやすい形で開発できるようにします。

![Demo](./images/demo.gif)

## xlflowとは

[xlflow](https://github.com/harumiWeb/xlflow) は、AIエージェントがExcel VBAマクロを自律的に開発可能にすることをファーストに作られた開発支援CLIツールです。

Excelブック内のVBAコードを .bas / .cls / .frm などのファイルとして取り出し、Gitで管理し、編集後に再びブックへ反映できます。

また、VBAマクロの実行、テスト、lint、format、静的解析などをCLIから実行できるため、人間による開発だけでなく、AIエージェントを使ったExcel VBA開発にも適しています。

xlflow for VSCode では**上記の機能群をGUI化し、LSPサーバーも提供することで人間にとっても非常に優れた開発体験を提供**します。

## 動作要件

- 対象OSは **Windows** のみです。
- `xlflow`をインストールする必要があります。
- `xlflow`をシステムパスに追加するか、`xlflow.path`環境変数に実行ファイルの絶対パスを設定してください。
- Excelの設定から「**VBA プロジェクト オブジェクト モデルへのアクセスを信頼する**」を有効にしてください。

![Trust Setting](./images/trust_setting.png)

### xlflow本体のインストールコマンド

- Quick Install

  ```bash
  irm https://harumiweb.github.io/xlflow/install.ps1 | iex
  ```

- WinGet

  ```bash
  winget install HarumiWeb.Xlflow
  ```

- Scoop

  ```bash
  scoop bucket add harumiweb https://github.com/harumiWeb/scoop-bucket
  scoop install xlflow
  ```

- WSLにインストールする場合（別途Windows側にもインストールが必要です）

  ```bash
  curl -fsSL https://harumiweb.github.io/xlflow/install.sh | sh
  ```

## この拡張機能でできること

xlflow for Visual Studio Code では、xlflow CLI の主要な操作を VSCode 上から実行できます。

主な機能は以下のとおりです。

- xlflowプロジェクトの状態表示
- `xlflow.toml` をもとにしたプロジェクト認識
- ExcelブックからVBAモジュールを取り込む
- 編集したVBAモジュールをExcelブックへ反映する
- xlflowセッションの開始・停止
- 自動テストの実行
- LSPによる入力補完とリアルタイムの診断
- ASTベースの静的解析とフォーマッター
- 標準モジュール・クラスモジュールなどの一覧表示
- コマンドパレットからのxlflowコマンド実行
- VBA開発向けの補助機能

## 想定している利用シーン

この拡張機能は、次のようなExcel VBA開発に向いています。

- Excel VBAマクロをGitで管理したい
- VBEではなくVSCodeでVBAコードを編集したい
- 既存のExcelマクロ資産を安全に保守したい
- VBAにもlintやformatを導入したい
- AIエージェントにExcel VBA開発を任せたい
- Excelブックとソースコードの同期操作をGUIから行いたい
- WSLやCLIベースの開発フローとExcel VBAをつなげたい

## 設定について

`xlflow`がシステムパスに登録されていない場合の実行ファイルパスを設定するには：

```json
{
  "xlflow.path": "C:\\path\\to\\xlflow.exe"
}
```

一般的な設定項目は以下の通りです：

- `xlflow.lsp.enabled`: VBAファイルに対して`xlflow lsp --stdio`を起動します。
- `xlflow.lsp.logFile`: 言語サーバに渡すログファイル。デフォルト値は`.xlflow/lsp.log`です。
- `xlflow.lsp.trace.server`: 言語サーバのトレース出力チャネルにおける詳細度設定。
- `xlflow.codeLens.enabled`: 実行可能なVBAプロシージャの上にxlflow CodeLensアクションを表示します。
- `xlflow.codeLens.runProcedure`: 実行可能なVBAプロシージャの上に`Run`（実行する）アクションを表示します。
- `xlflow.codeLens.runTests`: VBAテストプロシージャの上に`Run Test`（テストを実行する）アクションを表示します。
- `xlflow.codeLens.userFormEvents`: UserFormのイベントハンドラの上に`Run`（実行）アクションを表示します。
- `xlflow.run.saveBeforeRun`: CodeLensからプロシージャを実行する前に、変更済みのVBAドキュメントを保存します。
- `xlflow.completion.triggerSuggestInStatements`: VBAステートメントが記述される可能性が高い文脈で、VS Codeのサジェスト機能を起動させます。
- `xlflow.completion.progIdsInStrings`: `CreateObject("...")`および`GetObject("...")`文字列内において、VS Codeのサジェスト機能をトリガーします。
- `xlflow.testing.autoDiscover`: xlflowワークスペースが開かれた際に、VBAテストを自動的に検出します。

## コマンドについて

コマンドパレットには以下の機能が含まれています：

| コマンド | 説明 |
| --- | --- |
| `xlflow: Restart Language Server` | 補完、診断、ジャンプがずれたときに VBA Language Server を再起動します。 |
| `xlflow: Check Environment` | xlflow、Excel 連携、現在のワークスペースが利用可能か確認します。 |
| `xlflow: New Project` | 新しい xlflow プロジェクトのひな形を作成します。 |
| `xlflow: Initialize Project` | 既存のブックプロジェクトに xlflow 設定を追加します。 |
| `xlflow: Install Agent Skill` | xlflow 用の AI エージェント Skill をインストールします。 |
| `xlflow: Install Helper Modules` | xlflow の機能やサンプルで使う補助 VBA モジュールを追加します。 |
| `xlflow: New Module` | 種類を選んで新しい VBA モジュールを作成します。 |
| `xlflow: New Standard Module` | 新しい標準モジュールを作成します。 |
| `xlflow: New Class Module` | 新しいクラスモジュールを作成します。 |
| `xlflow: New UserForm` | 新しい UserForm 一式を作成します。 |
| `xlflow: Pull Workbook` | 現在のブックにある VBA 資産をワークスペースへ取り込みます。 |
| `xlflow: Push Sources` | ワークスペース上のソース変更をブックへ反映します。 |
| `xlflow: Run Macro` | 設定済みのエントリーマクロを実行します。 |
| `xlflow: Run Procedure` | 選択した VBA プロシージャを実行します。 |
| `xlflow: Run Test Procedure` | 選択した VBA テストプロシージャを直接実行します。 |
| `xlflow: Run Tests` | プロジェクトの VBA テスト一式を実行します。 |
| `xlflow: Lint Workspace` | ワークスペースのソースに対して lint を実行します。 |
| `xlflow: Format Document` | アクティブな VBA ドキュメントを整形します。 |
| `xlflow: Format Project` | プロジェクト内の対応ソースをまとめて整形します。 |
| `xlflow: Save Workbook` | 接続中の Excel ブックを保存します。 |
| `xlflow: Start Session` | 繰り返し実行を高速化する再利用可能な Excel セッションを開始します。 |
| `xlflow: Session Status` | 現在の xlflow セッション状態を表示します。 |
| `xlflow: Restart Session` | 管理中の Excel セッションを開き直します。 |
| `xlflow: Stop Session` | アクティブな xlflow セッションを停止します。 |
| `xlflow: Open Output` | VS Code の xlflow 出力チャネルを開きます。 |
| `xlflow: Refresh Project` | プロジェクトツリーと関連状態を再読み込みします。 |
| `xlflow: Refresh Modules` | サイドバーのモジュール一覧を更新します。 |
| `xlflow: Refresh UserForms` | サイドバーの UserForm 一覧を更新します。 |
| `xlflow: Refresh Tests` | テストエクスプローラーの検出済みテストを更新します。 |
| `xlflow: Run All Tests` | サイドバーやテストビューから検出済みの全 VBA テストを実行します。 |
| `xlflow: Run Doctor` | 詳細な環境診断のために `xlflow doctor` を実行します。 |
| `xlflow: Toggle Session` | 現在のワークスペースでセッションモードをオン・オフします。 |
| `xlflow: Open Documentation` | xlflow のドキュメントを開きます。 |
| `xlflow: Rename Module` | VBA モジュール名と対応するソースファイル名を変更します。 |
| `xlflow: Delete Module` | ワークスペースからモジュールを削除します。 |
| `xlflow: Reveal Source File` | 選択したモジュールソースの場所を開きます。 |
| `xlflow: Copy Module Name` | 選択したモジュール名をクリップボードへコピーします。 |
| `xlflow: Copy Relative Path` | 選択したソースファイルのプロジェクト相対パスをコピーします。 |
| `xlflow: Copy Procedure Name` | 選択したプロシージャ名をコピーします。 |
| `xlflow: Copy Qualified Name` | モジュール名を含む完全修飾プロシージャ名をコピーします。 |
| `xlflow: Rename UserForm` | UserForm 名と関連成果物の名前を変更します。 |
| `xlflow: Delete UserForm` | ワークスペースから UserForm を削除します。 |
| `xlflow: Reveal UserForm Source` | 選択した UserForm ソースの場所を開きます。 |
| `xlflow: Copy UserForm Name` | 選択した UserForm 名をコピーします。 |
| `xlflow: Copy UserForm Relative Path` | 選択した UserForm ソースのプロジェクト相対パスをコピーします。 |

## AIエージェントとの連携

xlflowは、AIエージェントにExcel VBAマクロを開発させるために作られたツールであり、すべての操作はCLIで操作可能でAIフレンドリーな構造化出力を備えています。

ターミナルから

```bash
xlflow skill install
```

またはVSCodeのコマンドパレットから

```bash
xlflow: Install Agent Skill
```

を使ってAIエージェント向けの **Agent Skill** をインストールすることで、Codex / Claude Code / GitHub Copilot / Cursor といったあらゆるコーディングエージェントがxlflowを巧みに使いこなし、完全自律してExcel VBAマクロを開発することが可能です。

これにより、Excel VBA開発にもテスト駆動開発や自動修正のワークフローを取り入れやすくなります。

![Ai-Driven Development](./images/ai-drive-develop.gif)

## WSL連携について

xlflowは、Windows上のExcelとWSL上の開発環境をつなぐワークフローに対応しています。

WSL上のエディタやAIエージェントからVBAコードを編集し、Windows側のExcelに対して取り込み・反映・実行を行うことができます。

WSL連携を利用する場合は、以下に注意してください。

Windows側とWSL側の両方にxlflowをインストールする必要があります
対象プロジェクトは `/mnt/c/...` 配下など、WindowsとWSLの両方から参照できる場所に置く必要があります
Excelブック操作にはWindows側のExcelが必要です

詳しくは [xlflow公式ドキュメント](https://harumiweb.github.io/xlflow/installation#wsl-development-frontend) を参照してください。

## トラブルシューティング

### xlflowコマンドが見つからない

`xlflow` CLI がインストールされているか確認してください。

ターミナルで以下を実行します。

```bash
xlflow version
```

コマンドが見つからない場合は、xlflow CLIをインストールするか、VSCode設定で `xlflow.path` を指定してください。

### プロジェクトとして認識されない

ワークスペース直下、または対象フォルダ内に `xlflow.toml` が存在するか確認してください。

```txt
my-project/
  xlflow.toml
```

xlflow.toml が存在しない場合は、コマンドパレットまたは専用のサイドバーからプロジェクト初期化を実行してください。

```bash
xlflow: Initialize Project
```

### Excelブック操作に失敗する

Excelブック操作には、Windows上のMicrosoft Excelが必要です。

以下を確認してください。

- Microsoft Excelがインストールされているか
- 対象ブックが開ける状態か
- VBAプロジェクトへのアクセスが許可されているか
- ブックが保護されていないか
- 別のExcelプロセスがブックをロックしていないか

### WSLから操作できない

WSL連携を利用する場合は、プロジェクトがWindows側からもWSL側からも参照できる場所にある必要があります。

推奨される配置例:

```txt
/mnt/c/dev/my-xlflow-project
```

また、Windows側とWSL側の両方で xlflow が実行できることを確認してください。

## 既知の制限事項

- この拡張機能では`xlflow`のインストールまたはバンドルは行いません。
- マクロ選択はまだインタラクティブな操作に対応していません。`xlflow: Run Macro`を実行すると、設定済みのデフォルトマクロが実行されます。引数なしで使える`Sub`プロシージャは、CodeLensから起動可能です。
- `xlflow: New Project`および`xlflow: Initialize Project`では、基本CLIワークフローのみが表示され、`--with-skill`、`--with-module`、`--agent`、`--json`オプションを選択するためのピッカーは提供されません。
- この拡張機能ではVBAコードの解析、診断、フォーマット、補完候補表示、シンボル分析、またはTypeScriptにおける型推論機能を実装していません。

## ドキュメント

詳しい使い方は、以下のドキュメントを参照してください。

[xlflow Documentation](https://harumiweb.github.io/xlflow/)
[GitHub Repository](https://github.com/harumiWeb/xlflow)

## フィードバック・不具合報告

不具合報告、機能要望、質問は GitHub Issues へお願いします。

[Issues](https://github.com/harumiWeb/xlflow/issues)

報告時には、可能であれば以下の情報を含めてください。

- 使用しているOS
- VSCodeのバージョン
- xlflowのバージョン
- この拡張機能のバージョン
- 実行したコマンド
- エラーメッセージ
- 再現手順

## 開発について

Node.js 22以降を使用してください。拡張機能のテストランナーには`@vscode/test-electron` 3.xを使用しています。
このディレクトリから：

```bash
pnpm install
pnpm compile
```

VS Codeの開発環境モードで拡張機能を起動するには、このフォルダを開き、コンパイル後に［実行とデバッグ］ビューから拡張機能ホストを実行してください。

## ライセンス

Mit License
