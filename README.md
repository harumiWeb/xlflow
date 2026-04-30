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

**An Excel VBA development framework for the AI agent era**

xlflow turns Excel VBA projects into a CLI-first development workflow.

Traditional Excel VBA development usually means editing code directly in the VBE, running macros manually, and debugging problems through the Excel UI. That workflow is cumbersome for humans and especially unsuitable for AI agents that operate primarily through the command line.

xlflow lets you manage VBA as normal source code and inspect, import, run, test, trace, and compare Excel VBA projects from the CLI.

## What xlflow can do

- Export VBA modules from `.xlsm` workbooks
- Edit `.bas`, `.cls`, and `.frm` files as normal source code
- Import edited VBA source back into Excel workbooks
- Run macros from the CLI
- Run VBA tests from the CLI
- Compare workbook cell values, formulas, and exported VBA source
- Lint VBA for automation-hostile patterns
- Detect GUI interaction boundaries before unattended runs
- Collect runtime logs through trace events
- Return stable JSON for AI agents
- Install bundled Skills for Codex, Claude, Cursor, Gemini, and other agents

## Why xlflow exists

Excel VBA is still an important automation platform in many business environments. At the same time, it is difficult for AI agents to work with safely.

Common problems include:

- VBA code is trapped inside `.xlsm` files
- Editing, running, and verifying VBA from the CLI is hard
- Macro entrypoints are often unclear
- Runtime errors are difficult to locate and diagnose
- File picker dialogs and `MsgBox` calls block unattended automation
- Tests and diffs are hard to standardize

xlflow provides the following development loop for Excel VBA:

```text
pull → edit → push → lint → test/run → trace → diff
```

This makes Excel VBA closer to normal software development for both humans and AI agents.

## Requirements

xlflow is Windows-first.

Excel operations use PowerShell and Excel COM. Commands that operate on workbooks, such as `pull`, `push`, `run`, `test`, `macros`, and `trace`, require Windows and Microsoft Excel.

Reading and writing VBA projects also requires enabling **Trust access to the VBA project object model** in Excel.

Commands that do not require Excel COM, such as `lint`, parts of `diff`, and Go unit tests, can be verified in non-Excel environments.

## Installation

If Go is available, install xlflow with:

```bash
go install github.com/harumiWeb/xlflow/cmd/xlflow@latest
```

After installation, verify that the command is available:

```bash
xlflow --help
```

To run xlflow directly from a development checkout:

```bash
go run ./cmd/xlflow --help
```

If you use Taskfile:

```bash
task run -- --help
```

## Quick start

Create a new xlflow project and macro-enabled workbook:

```bash
xlflow new Book.xlsm
```

To install the AI agent Skill during project creation:

```bash
xlflow new Book.xlsm --with-skill --agent codex
```

To start from an existing Excel workbook:

```bash
xlflow init Book.xlsm
```

Check Excel, COM, and VBIDE access:

```bash
xlflow doctor --json
```

Export VBA from the workbook into source files:

```bash
xlflow pull --json
```

After editing VBA source, import it back into the workbook:

```bash
xlflow push --json
```

Discover runnable macro entrypoints:

```bash
xlflow macros --json
```

Run a macro:

```bash
xlflow run Main.Run --json
```

For unattended automation, prefer headless mode:

```bash
xlflow run Main.Run --headless --json
```

If the macro intentionally shows file pickers, message boxes, or UserForms, use interactive mode so a person can operate Excel:

```bash
xlflow run Main.Run --interactive --timeout 5m --json
```

Run VBA tests:

```bash
xlflow test --json
```

Run the linter:

```bash
xlflow lint --json
```

## Commands

### `xlflow new`

Creates a new xlflow project and `.xlsm` workbook.

```bash
xlflow new
xlflow new Sales
xlflow new Sales.xlsm
```

When no argument is provided, xlflow creates `build/Book.xlsm`. When the name has no extension, `.xlsm` is appended. `new` creates a macro-enabled workbook, so extensions other than `.xlsm` are rejected.

`new` creates the project structure, including `xlflow.toml`, `src/`, `tests/`, `build/`, and `.xlflow/`. It also creates or updates `.gitignore` to ignore Excel temporary files and xlflow-generated artifacts.

### `xlflow init`

Creates an xlflow project from an existing Excel workbook.

```bash
xlflow init Book.xlsm
```

The given workbook is copied under `build/`, and its project-local path is recorded in `[excel].path` in `xlflow.toml`.

### `xlflow doctor`

Diagnoses the Excel automation environment.

```bash
xlflow doctor --json
```

It checks whether Excel is installed, whether the workbook can be opened, and whether VBIDE access is available. If `pull`, `push`, `run`, or `test` fails because of the environment, run `doctor` first.

When source files are available, `doctor` also reports GUI boundary candidates that may block headless runs.

### `xlflow attach`

Validates the workbook currently active in Excel.

```bash
xlflow attach --active --json
```

This is a safety check for human-assisted sessions. It confirms that the active Excel workbook matches configured `excel.path`; it does not change the target used by `pull`, `push`, or `run`.

### `xlflow pull`

Exports VBA components from the configured workbook.

```bash
xlflow pull --json
```

It exports standard modules, class modules, UserForms, and document modules such as Workbook and Worksheet modules into `src/`.

### `xlflow push`

Imports VBA source under `src/` back into the Excel workbook.

```bash
xlflow push --json
```

It reads `.bas`, `.cls`, and `.frm` files and imports them through VBIDE. UserForm `.frx` files are treated as binary companion files.

### `xlflow macros`

Discovers runnable Public Sub entrypoints.

```bash
xlflow macros --json
```

AI agents and automation scripts should run this command before guessing a macro name. Use the returned `qualified_name` with `xlflow run` to avoid entrypoint mistakes.

### `xlflow run`

Runs a macro from the CLI.

```bash
xlflow run Main.Run --json
```

Macros with arguments are supported:

```bash
xlflow run Report.Generate \
  --arg string:fixtures\sample.xlsx \
  --arg int:3 \
  --arg bool:true \
  --json
```

`--arg` accepts typed arguments with `string:`, `int:`, and `bool:` prefixes. Empty values are allowed only for `string:`.

By default, `run` does not save the workbook. To persist results, explicitly pass `--save` or `--save-as`.

```bash
xlflow run Report.Generate --save --json
xlflow run Report.Generate --save-as build\Result.xlsm --json
```

When execution fails, xlflow returns `macro_failed` or `macro_not_found` with VBA error number, description, module name, phase, and line number when available.

`--headless` rejects GUI boundaries before Excel starts and returns `gui_boundary_detected` with top-level `gui_boundaries`. `--interactive` runs with Excel visible and alerts enabled for human operation. `--timeout` defaults to `5m` and returns `macro_timeout` when execution does not complete in time.

### `xlflow trace`

Collects log events from VBA during macro execution.

First, inject the trace module:

```bash
xlflow trace inject --json
```

Then write logs from VBA:

```vb
Call XlflowLog("start GenerateReport")
Call XlflowLog("lastRow=" & lastRow)
Call XlflowLog("finished GenerateReport")
```

Run the macro with trace enabled:

```bash
xlflow run Main.Run --trace --json
```

Trace events are returned in the top-level JSON `trace` field. This helps identify how far execution progressed before a runtime error.

### `xlflow test`

Runs VBA tests.

```bash
xlflow test --json
```

xlflow discovers argument-free `Sub` procedures whose names start with `Test` or end with `_Test`.

To run a single test, use `--filter`:

```bash
xlflow test --filter TestCreateReport --json
```

New and initialized projects include `src/modules/XlflowAssert.bas`. Use `AssertEquals expected, actual, [message]` to compare scalar values.

```vb
Public Sub TestCreateReport()
    AssertEquals 10, Sheets("Result").Range("A1").Value2
End Sub
```

`AssertEquals` does not support object or array comparison. Compare scalar properties such as `Range.Value2` instead of passing `Range` objects directly.

### `xlflow diff`

Compares two workbooks.

```bash
xlflow diff before.xlsm after.xlsm --json
```

It detects sheet additions/removals, cell value differences, and formula differences.

To compare exported VBA source as well, pass `--vba-before` and `--vba-after`:

```bash
xlflow diff before.xlsm after.xlsm \
  --vba-before before-src \
  --vba-after after-src \
  --json
```

Differences are reported as successful command results. Inspect `diff.summary.total_diffs` in JSON to determine whether anything changed.

### `xlflow lint`

Lints VBA source.

```bash
xlflow lint --json
```

It detects patterns that are unsafe or inconvenient for AI agents and unattended automation, including:

- Missing `Option Explicit`
- `Select` usage
- `Activate` usage
- `On Error Resume Next` usage
- Possible implicit `Variant`
- Module-level `Public` variables
- Interactive operations such as `Application.GetOpenFilename`, `Application.FileDialog`, `InputBox`, and modal `MsgBox`

### `xlflow inspect-gui`

Reports GUI interaction boundaries without opening Excel.

```bash
xlflow inspect-gui --json
```

The report includes file, line, kind, symbol, and a suggested refactor. Use it before deciding whether a macro should run with `--headless` or `--interactive`.

### `xlflow skill install`

Installs the bundled xlflow Skill for AI agents.

```bash
xlflow skill install --agent codex
xlflow skill install --agent claude
xlflow skill install --agent cursor
xlflow skill install --agent gemini
xlflow skill install --target .agents/skills
```

Supported provider targets are:

- `agents`: `.agents/skills/xlflow`
- `codex`: `.codex/skills/xlflow`
- `claude`: `.claude/skills/xlflow`
- `cursor`: `.cursor/skills/xlflow`
- `gemini`: `.gemini/skills/xlflow`

For GitHub Copilot, use the shared `.agents` target:

```bash
xlflow skill install --agent agents
```

## Configuration

xlflow reads `xlflow.toml` from the project root.

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

`project.entry` is used when `xlflow run` is invoked without a macro name.

## JSON output

Every command can return AI-agent-friendly JSON by passing `--json`.

The basic envelope is:

```json
{
  "status": "ok",
  "command": "lint",
  "error": null,
  "logs": []
}
```

On failure, `status` is `failed`, and `error.code` and `error.message` are returned.

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

AI agents and CI jobs should use `--json` instead of parsing human-readable output.

## Exit codes

xlflow uses the following exit code categories:

| Code | Meaning                                                             |
| ---: | ------------------------------------------------------------------- |
|    0 | Success                                                             |
|    1 | Validation failure, such as lint, macro, or test failure            |
|    2 | CLI argument or configuration error                                 |
|    3 | Environment error, such as Excel, COM, VBIDE, or PowerShell failure |

`diff` returns exit code `0` even when differences are found. Inspect `diff.summary.total_diffs` to determine whether inputs differ.

## Using xlflow with AI agents

xlflow provides a proof loop for AI agents to safely edit, run, and verify Excel VBA.

Recommended workflow:

```text
1. Read xlflow.toml
2. If needed, run xlflow pull --json to refresh current VBA source
3. Edit .bas, .cls, or .frm files under src/
4. Run xlflow push --json to update the workbook
5. Run xlflow lint --json and fix unsafe patterns
6. Run xlflow test --json
7. If no tests exist, run xlflow macros --json → xlflow run <qualified_name> --headless --json
8. If runtime errors are unclear, use xlflow trace inject → xlflow run --trace --json
9. If workbook changes must be reviewed, use xlflow diff --json
```

To give the workflow to an AI agent, install the bundled Skill with `xlflow skill install` or `xlflow new/init --with-skill`.

```bash
xlflow skill install --agent codex
```

## Recommended VBA rules

VBA executed by xlflow should be written for unattended automation.

- Always use `Option Explicit`
- Do not rely on `Select`, `Activate`, or `ActiveSheet`
- Use explicit `Workbook`, `Worksheet`, and `Range` references
- Prefer `Long` over `Integer`
- Do not depend on UI dialogs or modal `MsgBox`
- Keep GUI entrypoints thin and extract parameterized headless procedures for the core logic
- Pass input values through `xlflow run --arg`, configuration files, deterministic paths, or environment variables
- Avoid broad `On Error Resume Next`
- Emit error messages that make failures diagnosable
- Verify destructive workbook changes with tests or diff

## Local verification

Run the fast repository verification with:

```bash
task verify
```

Currently, `task verify` runs `go test ./...` as non-COM test coverage.

Excel COM E2E verification should be run on Windows with Excel and VBIDE access enabled.

```bash
xlflow doctor --json
```

After `doctor` reports a healthy environment, run `new`, `doctor`, `pull`, `lint`, `push`, `run`, `test`, and `diff` against a real workbook.

## Current status

xlflow is an MVP-stage tool.

Its primary goal is to bring Excel VBA into AI-agent and CLI-based development workflows. Typical use cases include:

- Source control for existing VBA
- AI-agent-assisted VBA modification
- CLI execution of Excel macros
- Automated VBA testing
- Debugging with runtime logs
- Workbook change review through diff
- More maintainable internal Excel automation

## License

MIT
