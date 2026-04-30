# xlflow CLI Contract

## Scope

This spec defines the MVP command, configuration, JSON output, and exit-code contracts for xlflow.

xlflow is a Windows-first Go CLI that treats Excel VBA projects as source-controlled code. Excel operations use PowerShell and Excel COM. Non-Excel commands such as `init` and `lint` should remain testable without Excel installed.

## Commands

```text
xlflow [--json] new [workbook] [--with-skill] [--agent <provider>]
xlflow [--json] init <workbook> [--with-skill] [--agent <provider>]
xlflow [--json] doctor
xlflow [--json] pull
xlflow [--json] push
xlflow [--json] trace inject [workbook]
xlflow [--json] run [macro] [--input <workbook>] [--arg <type:value>]... [--save | --save-as <path>] [--trace]
xlflow [--json] macros
xlflow [--json] test [--filter <name>]
xlflow [--json] diff <before-workbook> <after-workbook> [--vba-before <dir>] [--vba-after <dir>]
xlflow [--json] lint
xlflow [--json] skill install [--agent <provider> | --target <dir>] [--force]
```

`--json` is a persistent global flag and can be used with every command, including `new` and `init`.

When `--json` is not set, output is optimized for humans rather than machines. Interactive terminals may use Bubble Tea/Lipgloss presentation, color, and progress spinners for Excel COM-backed commands. Non-interactive output, such as CI logs and pipes, stays static and text-oriented while preserving the same command result information. Machine consumers must use `--json` instead of parsing human output.

`new` creates a fresh macro-enabled workbook under `build/` and scaffolds the same project layout as `init`. Without an argument it creates `build/Book.xlsm`; when the argument has no extension, `.xlsm` is appended. Any other extension is rejected because workbook creation always uses Excel macro-enabled format `52`.

`init` accepts an existing workbook path, copies that workbook into the new project's `build/<basename>` path, and records that project-local `build/...` path in `xlflow.toml` under `[excel].path` (for example `build/Sales.xlsx`).

`new` and `init` create or update a project-local `.gitignore`. The managed entries ignore Excel temporary files (`~$*.xls*`, `*.tmp`) and xlflow-generated state (`.xlflow/`, `build/`). Existing `.gitignore` content is preserved; missing managed entries are appended without duplicating entries that are already present.

`new` and `init` do not create `prompts/agent.md`. Use `--with-skill` to install the bundled `xlflow` AI agent skill during project creation. `--agent` selects one of `agents`, `codex`, `claude`, `cursor`, or `gemini`. When `--with-skill` is used without `--agent` in an interactive terminal, xlflow opens a Bubble Tea provider selector. With `--json` or non-interactive input, `--agent` is required.

`skill install` installs the bundled `xlflow` skill without creating or changing an xlflow project scaffold. Provider targets are:

- `agents`: `.agents/skills/xlflow`
- `codex`: `.codex/skills/xlflow`
- `claude`: `.claude/skills/xlflow`
- `cursor`: `.cursor/skills/xlflow`
- `gemini`: `.gemini/skills/xlflow`

For GitHub Copilot, use `agents` because Copilot reads repository instructions from `.agents`. `--target <dir>` installs to `<dir>/xlflow` instead of a provider default. `--agent` and `--target` cannot be combined. Existing skill directories are not overwritten unless `--force` is set. If neither `--agent` nor `--target` is provided, interactive terminals use the Bubble Tea provider selector; `--json` and non-interactive runs return a configuration error instead.

`pull` exports standard modules, class modules, userforms, and workbook document modules into the configured source directories. Userforms may emit both `.frm` and `.frx` artifacts. Document modules are exported as source text suitable for linting and re-import. Source-controlled `.bas`, `.cls`, and `.frm` files are UTF-8 without BOM. Excel/VBIDE import and export files are treated as CP932 at the bridge boundary, and `pull` converts exported text to UTF-8 before writing the source tree.

`push` reads source-controlled `.bas`, `.cls`, and `.frm` files as UTF-8 without BOM, writes CP932 temporary import copies under `.xlflow/tmp/`, and imports those temporary files through VBIDE. `.frx` files are binary userform companions and are copied without text conversion.

`trace inject` injects or replaces the standard module `XlflowTrace` in the target workbook. When `[workbook]` is omitted, it uses `excel.path` from `xlflow.toml` and also writes the same bundled trace module source to `<src.modules>/XlflowTrace.bas` as UTF-8 without BOM. This keeps a subsequent `push` from deleting the workbook trace module. JSON output for configured project injection includes top-level `source.path` and `source.updated` metadata. When `[workbook]` is provided and project configuration cannot be loaded, the command can run without source persistence. The injected module provides `XlflowLog message` for user VBA code and `XlflowSetTraceFile path` for the run harness. `new` and `init` do not create this module by default because trace logging is opt-in debug instrumentation.

`run` uses the positional macro argument when provided. Otherwise it uses `project.entry` from `xlflow.toml`. `--input` overrides `excel.path` for one invocation. `--arg` may be repeated and must use explicit prefixes: `string:hello`, `string:`, `int:7`, and `bool:true`. Empty values are valid only for `string:` arguments. Malformed `int:` and `bool:` values are rejected by the CLI before Excel starts and exit with code `2`. The default run never saves. `--save` persists the opened workbook in place after a successful run. `--save-as` writes a copy after a successful run and must keep the same workbook extension as the opened workbook. `--save` and `--save-as` cannot be combined.

`run` adds a `macro` object with `name`, `args`, and `duration_ms`. Failed macro runs return `macro_failed` with `error.source`, `error.number`, `error.message`, `error.line` when VBA exposes a non-zero `Erl` value, and `error.phase` when the failed phase is known. Stable run phases are `open_workbook`, `prepare_vbide`, `inject_harness`, `invoke_macro`, `save_result`, and `read_trace`. When Excel exposes enough information to distinguish a missing or invalid target macro from user-code failure, `run` returns `macro_not_found` instead of `macro_failed`. Plain-text success output must include the elapsed duration and whether the workbook was saved, copied, or left unchanged. Plain-text failure output must use the formatted message `Module line <n> Err <n>: <description>` when line and error number are available, and otherwise omit the `line <n>` segment. Because `run` injects a temporary VBA harness to measure duration while avoiding modal VBA runtime error dialogs, VBIDE access failures return an environment error such as `vbide_access_denied` and exit code `3`.

`run --trace` creates a fresh temp log under `%TEMP%\xlflow`, calls `XlflowTrace.XlflowSetTraceFile` before the target macro, then reads trace events after execution. User VBA code writes events with `Call XlflowLog("message")`. JSON output adds top-level `trace.enabled`, `trace.path`, `trace.events`, and optional `trace.read_error`; each event has `timestamp`, `message`, and `raw`. Plain-text output prints trace events as `[timestamp] message` after the normal run logs. If `XlflowTrace` is missing, `run --trace` returns `trace_not_injected` with exit code `1`. If the macro fails after writing trace events, those events are still returned. Trace read errors are reported in `trace.read_error` without changing the macro result. If a traced run fails with zero events, output should indicate that execution may have failed before reaching user trace calls.

`macros` opens the configured workbook and discovers public runnable VBA entrypoints without executing user code. JSON output includes top-level `macros`, where each entry contains `module`, `name`, `qualified_name`, `kind` when available, and `args` when available. Agents should use this command before guessing a `run` target.

`test` opens the configured workbook, discovers argument-free `Sub` procedures from the workbook VBIDE state, and runs procedures whose names start with `Test` or end with `_Test`. `--filter` uses exact procedure-name matching. Duplicate discovered test names, no discovered tests, missing filter targets, and VBA test failures are validation failures. Excel, COM, VBIDE, PowerShell, and script failures are environment failures.

`diff` compares two workbook files and optionally two exported VBA source trees. Workbook inputs must use `.xlsx`, `.xlsm`, `.xltx`, or `.xltm`. Workbook state comparison covers sheet additions/removals plus used-range cell values and formulas. VBA comparison is enabled only when both `--vba-before` and `--vba-after` are provided, recursively compares `.bas`, `.cls`, and `.frm` files, ignores other files such as `.frx`, and normalizes CRLF/LF line endings before comparison. Differences are successful command results with exit code `0`; malformed arguments fail with exit code `2`, and unreadable workbooks or source trees fail with exit code `3`.

## Configuration

The MVP only auto-discovers `xlflow.toml` from the current working directory. `vba.toml` is intentionally not supported.

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

## JSON Envelope

All JSON output uses a stable top-level envelope.

```json
{
  "status": "ok",
  "command": "lint",
  "error": null,
  "logs": []
}
```

`status` is either `ok` or `failed`. `error` is `null` on success and a structured object on failure.

Command-specific fields are added at the top level:

- `diagnostics` for `doctor`
- `workbook` and `backup` for Excel file commands
- `source` for commands that write project source files
- `macro` for `run`
- `macros` for `macros`
- `tests` for `test`
- `diff` for `diff`
- `issues` for `lint`
- `trace` for traced `run`

`test` result objects contain `name`, `module`, `status`, `duration_ms`, and an optional `error`.

`diff` result objects contain `summary`, `sheets`, `cells`, and `vba`. Cell diffs contain `sheet`, `address`, `kind`, `before`, and `after`, where `kind` is `value` or `formula`. VBA diffs contain `file`, `kind`, and optional changed line details.

## Exit Codes

- `0`: success
- `1`: user-code or validation failure, including lint findings, macro failure, missing macro target, missing trace module, VBA test failure, no tests found, missing filter targets, and duplicate test names
- `2`: CLI argument or configuration error
- `3`: environment failure, including Excel, COM, VBIDE, PowerShell, and script execution failures

`diff` intentionally returns `0` when differences are found. Consumers should inspect `diff.summary.total_diffs` to distinguish changed and unchanged inputs.

## VBA Test Rules

New and initialized projects include `src/modules/XlflowAssert.bas` with `AssertEquals expected, actual, [message]`. The helper is scalar-only: it compares normal scalar values, treats `Null` as equal only to `Null`, and raises a clear assertion error for object or array inputs. Compare object properties such as `Range.Value2` instead of passing object references.

Example:

```vb
Public Sub TestCreateReport()
    AssertEquals 10, Sheets("Result").Range("A1").Value
End Sub
```

## Lint Rules

- `VB001`: missing `Option Explicit`
- `VB002`: `Select` usage
- `VB003`: `Activate` usage
- `VB004`: `On Error Resume Next` usage
- `VB005`: possible implicit `Variant`
- `VB006`: module-level `Public` variable usage
- `VB007`: automation-hostile interactive input such as `Application.GetOpenFilename`, `Application.FileDialog`, `InputBox`, or modal `MsgBox`
