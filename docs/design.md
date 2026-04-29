# 全体コンセプト

```text
xlflow = Agent-ready VBA development framework

「Excelを操作するCLI」ではなく
「VBAプロジェクトをコードとして扱う開発基盤」
```

---

# アーキテクチャ

```text
+---------------------------+
|        xlflow CLI (Go)      |
|---------------------------|
| command router            |
| config loader             |
| logger / JSON output      |
| linter / parser           |
| test runner               |
+-------------+-------------+
              |
              v
+---------------------------+
|   Execution Adapter Layer |
|---------------------------|
| PowerShell bridge         |
| VBScript bridge           |
+-------------+-------------+
              |
              v
+---------------------------+
|       Excel (COM)         |
| VBIDE / Application.Run   |
+---------------------------+
```

---

# コマンド設計（コア）

CLIは「人間」ではなく**AIが使う前提**で設計します。

## 基本コマンド

```bash
xlflow init Book.xlsm
xlflow pull
xlflow push
xlflow run Module1.Main
xlflow lint
xlflow doctor
```

---

## AI向け拡張コマンド（重要）

```bash
xlflow run --json
xlflow lint --json
xlflow diagnose --json
xlflow fix --auto
```

👉 全てJSONで返せるのが重要

---

# ディレクトリ構成

```text
project/
├─ src/
│  ├─ modules/
│  │   └─ Main.bas
│  ├─ classes/
│  │   └─ User.cls
│  ├─ forms/
│  │   └─ Form1.frm
│  └─ workbook/
│      └─ Sheet1.bas
│
├─ tests/
│  └─ MainTest.bas
│
├─ build/
│  └─ Book.xlsm
│
├─ prompts/              # AIエージェント用
│  └─ agent.md
│
├─ vba.toml
└─ .xlflow/
   └─ cache.json
```

---

# 設定ファイル（vba.toml）

```toml
[project]
name = "sales_tool"
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"
visible = false

[lint]
require_option_explicit = true
forbid_select = true
max_line_length = 120

[test]
runner = "xlflow"
```

---

# コア機能設計

## 1. pull / push

```text
pull
- VBComponents.Export
- ファイルとして保存

push
- 既存削除
- Import
```

---

## 2. run（最重要）

```bash
xlflow run Main.Run --json
```

出力：

```json
{
  "status": "failed",
  "macro": "Main.Run",
  "error": {
    "number": 91,
    "message": "Object variable not set",
    "module": "Main",
    "line": 42
  },
  "logs": [
    "Start processing...",
    "Loading sheet..."
  ]
}
```

👉 AIがこのJSONを元に修正できる

---

## 3. lint

VBA特有のルールを持つ

```text
- Option Explicit必須
- Select / Activate禁止
- 未使用変数
- 暗黙Variant
- Range("A1")直書き検出
- On Error Resume Next検出
```

---

## 4. doctor（かなり重要）

```bash
xlflow doctor
```

```json
{
  "excel_installed": true,
  "vbide_access": false,
  "fix": "Enable 'Trust access to VBA project model'"
}
```

👉 初期ハマりを完全に潰す

---

# AIエージェント対応設計

ここがこのツールの“核”です。

## prompts/agent.md

```md
You are a VBA developer.

Rules:
- Never use Select/Activate
- Always use Option Explicit
- Prefer With blocks
- Avoid global state
```

---

## AIフレンドリー設計

```text
すべてのコマンドが
- exit code
- JSON
- deterministic output
を持つ
```

---

# 実行ブリッジ設計

## PowerShell（推奨）

```powershell
$excel = New-Object -ComObject Excel.Application
$wb = $excel.Workbooks.Open($path)
$excel.Run("Main.Run")
```

---

## エラー取得

```text
On Error GoTo Handler
↓
Err.Number / Err.Description を取得
```

👉 JSONでGo側へ返す

---

# 将来拡張

## 1. テストフレームワーク

```vba
Sub Test_Add()
    AssertEqual Add(1,2), 3
End Sub
```

```bash
xlflow test
```

---

## 2. watchモード

```bash
xlflow watch
```

```text
保存 → push → run → 結果表示
```

👉 AIエージェントと相性抜群

---

## 3. スナップショット

```bash
xlflow snapshot save
xlflow snapshot restore
```

👉 Excel破壊対策

---

## 4. GUIログビュー（将来）

CLI + TUI（Bubble Tea）もあり

---

# MVPスコープ

まずここだけで十分価値があります

```text
- init
- pull / push
- run（JSON出力）
- lint（最低限）
- doctor
```

---

以下、**Go製CLI + PowerShellブリッジ**前提の初期設計です。

# ディレクトリ構成

```text
xlflow/
├─ cmd/
│  └─ xlflow/
│     └─ main.go
├─ internal/
│  ├─ cli/
│  │  └─ root.go
│  ├─ command/
│  │  ├─ init.go
│  │  ├─ pull.go
│  │  ├─ push.go
│  │  ├─ run.go
│  │  ├─ lint.go
│  │  └─ doctor.go
│  ├─ config/
│  │  └─ config.go
│  ├─ excel/
│  │  ├─ bridge.go
│  │  ├─ powershell.go
│  │  └─ result.go
│  ├─ lint/
│  │  ├─ linter.go
│  │  └─ rules.go
│  ├─ project/
│  │  ├─ layout.go
│  │  └─ scaffold.go
│  └─ output/
│     └─ json.go
├─ scripts/
│  ├─ pull.ps1
│  ├─ push.ps1
│  ├─ run.ps1
│  └─ doctor.ps1
├─ testdata/
│  └─ sample.xlsm
├─ go.mod
├─ README.md
└─ vba.toml
```

# Goパッケージ責務

```text
cmd/xlflow
- エントリーポイント

internal/cli
- CobraなどのCLIルート定義

internal/command
- 各コマンドのUseCase層

internal/config
- vba.toml読み込み

internal/excel
- PowerShell実行
- JSON結果のパース
- Excel操作の抽象化

internal/lint
- VBAファイル解析
- ルール実行

internal/project
- init時の雛形生成
- src/ tests/ build/ prompts/ 作成

internal/output
- 人間向け/AI向け出力の切り替え
```

# CLI設計

```bash
xlflow init ./Book.xlsm
xlflow pull
xlflow push
xlflow run
xlflow run Main.Run
xlflow lint
xlflow doctor
```

AI向けには共通で `--json` を持たせます。

```bash
xlflow run Main.Run --json
xlflow lint --json
xlflow doctor --json
```

# vba.toml

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
```

# PowerShellブリッジ方針

GoからPowerShellを実行します。

```go
cmd := exec.Command(
    "powershell",
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", scriptPath,
    "-WorkbookPath", workbookPath,
)
```

PowerShell側は必ずJSONを返します。

```json
{
  "status": "ok",
  "message": "Macro executed successfully"
}
```

失敗時：

```json
{
  "status": "failed",
  "error": {
    "number": 1004,
    "message": "Cannot run the macro",
    "source": "Microsoft Excel"
  }
}
```

# run.ps1 の最小イメージ

```powershell
param(
  [string]$WorkbookPath,
  [string]$MacroName,
  [bool]$Visible = $false
)

$result = @{
  status = "ok"
  macro = $MacroName
  error = $null
}

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $Visible
  $excel.DisplayAlerts = $false

  $workbook = $excel.Workbooks.Open($WorkbookPath)
  $excel.Run($MacroName)

  $workbook.Save()
  $workbook.Close($true)
  $excel.Quit()
}
catch {
  $result.status = "failed"
  $result.error = @{
    message = $_.Exception.Message
    source = $_.Exception.Source
  }

  if ($workbook) {
    $workbook.Close($false)
  }
  if ($excel) {
    $excel.Quit()
  }
}
finally {
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

$result | ConvertTo-Json -Depth 5
```

# lint MVPルール

まずは文字列ベースで十分です。

```text
VB001 Option Explicit がない
VB002 Select を使用している
VB003 Activate を使用している
VB004 On Error Resume Next を使用している
VB005 暗黙Variantの可能性がある
VB006 Public変数を使用している
```

出力例：

```json
{
  "status": "failed",
  "issues": [
    {
      "code": "VB002",
      "severity": "warning",
      "file": "src/modules/Main.bas",
      "line": 12,
      "message": "Avoid Select. Use direct object references instead."
    }
  ]
}
```

# MVP実装順

```text
1. xlflow init
2. vba.toml読み込み
3. xlflow doctor
4. xlflow pull
5. xlflow push
6. xlflow run --json
7. xlflow lint
```

最初に作るなら、`doctor` を先に作るのが良いです。
Excelが入っているか、COMが使えるか、VBIDEアクセスが許可されているかを確認できないと、他の機能のデバッグが地獄になります。

# 最初のREADME見出し

```md
# xlflow

Agent-ready VBA development framework.

xlflow turns Excel VBA projects into a CLI-first development workflow.

- Export VBA modules from `.xlsm`
- Edit VBA as normal source files
- Import modules back into Excel
- Run macros from CLI
- Lint VBA for safer automation
- Return machine-readable JSON for AI agents
```

この方針なら、かなり実装に移しやすいです。


---

# 総評（重要）

このプロジェクトの本質は

```text
VBAをCLI化することではない
↓
AIがExcel業務を自動化できるようにすること
```

です。
