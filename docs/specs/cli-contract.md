# xlflow CLI Contract

## Scope

This spec defines the MVP command, configuration, JSON output, and exit-code contracts for xlflow.

xlflow is a Windows-first Go CLI that treats Excel VBA projects as source-controlled code. Excel operations use PowerShell and Excel COM. Non-Excel commands such as `init` and `lint` should remain testable without Excel installed.

## Commands

```text
xlflow [--json] new [workbook] [--with-skill] [--agent <provider>] [--keepalive] [--keepalive-interval <duration>]
xlflow [--json] init <workbook> [--with-skill] [--agent <provider>]
xlflow [--json] doctor [--keepalive] [--keepalive-interval <duration>]
xlflow [--json] attach --active [--keepalive] [--keepalive-interval <duration>]
xlflow [--json] pull [--session] [--keepalive] [--keepalive-interval <duration>]
xlflow [--json] push [--backup always|never] [--fast] [--changed-only] [--session] [--no-save] [--keepalive] [--keepalive-interval <duration>]
xlflow [--json] session start
xlflow [--json] session status
xlflow [--json] session stop
xlflow [--json] save --session
xlflow [--json] runner install
xlflow [--json] runner remove
xlflow [--json] runner status
xlflow [--json] trace enable [workbook] [--session] [--keepalive] [--keepalive-interval <duration>]
xlflow [--json] trace disable [workbook] [--force] [--session] [--keepalive] [--keepalive-interval <duration>]
xlflow [--json] trace status [workbook] [--session] [--keepalive] [--keepalive-interval <duration>]
xlflow [--json] trace clean [--keepalive] [--keepalive-interval <duration>]
xlflow [--json] trace inject [workbook] [--session] [--keepalive] [--keepalive-interval <duration>]
xlflow [--json] run [macro] [--input <workbook>] [--arg <type:value>]... [--save | --save-as <path>] [--trace] [--headless | --interactive] [--direct] [--fast] [--session] [--timeout <duration>] [--keepalive] [--keepalive-interval <duration>]
xlflow [--json] macros [--session] [--keepalive] [--keepalive-interval <duration>]
xlflow [--json] ui button add --sheet <name> --cell <A1> --text <caption> --macro <module.proc> [--id <id>] [--width <points>] [--height <points>] [--create-sheet] [--verify-macro] [--keepalive] [--keepalive-interval <duration>]
xlflow [--json] ui button list [--sheet <name>] [--keepalive] [--keepalive-interval <duration>]
xlflow [--json] ui button remove --id <id> [--sheet <name>] [--keepalive] [--keepalive-interval <duration>]
xlflow [--json] test [--filter <name>] [--session] [--keepalive] [--keepalive-interval <duration>]
xlflow [--json] diff <before-workbook> <after-workbook> [--vba-before <dir>] [--vba-after <dir>]
xlflow [--json] inspect-gui
xlflow [--json] lint
xlflow [--json] analyze
xlflow [--json] check [--keepalive] [--keepalive-interval <duration>]
xlflow [--json] skill install [--agent <provider> | --target <dir>] [--force]
```

`--json` is a persistent global flag and can be used with every command, including `new` and `init`.

When `--json` is not set, output is optimized for humans rather than machines. Interactive terminals may use Bubble Tea/Lipgloss presentation, color, and progress spinners for Excel COM-backed commands. Non-interactive output, such as CI logs and pipes, stays static and text-oriented while preserving the same command result information. Machine consumers must use `--json` instead of parsing human output.

Excel COM-backed commands support `--keepalive` for AI agent and task-runner environments that may treat long silent Excel COM operations as stalled. This includes `new`, `doctor`, `attach`, `pull`, `push`, `trace enable/disable/status/inject/clean`, `run`, `macros`, `ui button add/list/remove`, `test`, and `check`. When enabled, xlflow writes heartbeat lines to stderr while the PowerShell/Excel bridge is still running, starting immediately and then repeating every `--keepalive-interval` duration. The default interval is `5s`; non-positive intervals are CLI argument errors when keepalive is enabled. Keepalive output never writes to stdout, so `--json` stdout remains a single machine-readable envelope. At completion, xlflow writes a stderr marker such as `XLFLOW_DONE status=success command=pull` or `XLFLOW_DONE status=failed command=run code=macro_timeout`. Agents should not begin the next workbook-dependent step until the command exits and this marker has been observed.

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

`pull` exports standard modules, class modules, userforms, and workbook document modules into the configured source directories. Userforms may emit both `.frm` and `.frx` artifacts. Document modules are exported as source text suitable for linting and re-import. Source-controlled `.bas`, `.cls`, and `.frm` files are UTF-8 without BOM. Excel/VBIDE import and export files are treated as CP932 at the bridge boundary, and `pull` converts exported text to UTF-8 before writing the source tree. `pull --session` exports from the workbook opened by `session start`.

`push` reads source-controlled `.bas`, `.cls`, and `.frm` files as UTF-8 without BOM, writes CP932 temporary import copies under `.xlflow/tmp/`, and imports those temporary files through VBIDE. `.frx` files are binary userform companions and are copied without text conversion. By default `push` creates a timestamped backup under `.xlflow/backups`, replaces non-document VBA components, updates document modules, saves the workbook, and writes source fingerprints to `.xlflow/state/push.json`.

`push --backup=never` skips the export backup. `push --fast` is a development-mode shorthand for `--backup=never --changed-only`. `push --changed-only` compares source fingerprints against `.xlflow/state/push.json`; when unchanged, it skips Excel/VBIDE import and returns `source.changed=false`. When changed or state is missing, v1 safely falls back to the normal full component replacement and refreshes the state file after success. `push --session` attaches to the workbook kept open by `xlflow session start` instead of opening a fresh Excel instance. `push --no-save` is allowed only with `--session` and leaves workbook changes unsaved until `xlflow save --session` or `xlflow session stop`.

`session start` opens the configured workbook in Excel and writes `.xlflow/session.json` with process metadata. `session status` reports whether the recorded process is running and the configured workbook is open. `session stop` saves and closes the workbook, quits Excel, and removes the metadata. Session v1 is explicit opt-in: normal `push` and `run` do not auto-attach to sessions. `save --session` saves the workbook held by the active session.

`runner install`, `runner remove`, and `runner status` manage the persistent workbook module `XlflowRunner`. In v1 this module is a stable marker for fast-run workflows; argument-free `run --fast` uses direct execution when eligible and otherwise keeps the normal temporary harness path.

`trace enable` injects or replaces the standard module `XlflowTrace` in the target workbook. When `[workbook]` is omitted, it uses `excel.path` from `xlflow.toml` and also writes the same bundled trace module source to `<src.modules>/XlflowTrace.bas` as UTF-8 without BOM. This keeps a subsequent `push` from deleting the workbook trace module. JSON output for configured project injection includes top-level `source.path` and `source.updated` metadata. `trace inject` is a compatibility alias for `trace enable`. `trace disable` removes the workbook helper and removes source helper only when it matches xlflow's bundled helper, unless `--force` is set. `trace status` reports workbook and source helper presence plus whether the source matches the bundled helper. `trace clean` removes `.xlflow/traces`. `trace enable/disable/status/inject --session` operate on the workbook opened by `session start`. The injected module provides `XlflowLog message` for user VBA code and `XlflowSetTraceFile path` for the run harness. `new` and `init` do not create this module by default because trace logging is opt-in debug instrumentation.

`run` uses the positional macro argument when provided. Otherwise it uses `project.entry` from `xlflow.toml`. `--input` overrides `excel.path` for one invocation. `--arg` may be repeated and must use explicit prefixes: `string:hello`, `string:`, `int:7`, and `bool:true`. Empty values are valid only for `string:` arguments. Malformed `int:` and `bool:` values are rejected by the CLI before Excel starts and exit with code `2`. The default run never saves. `--save` persists the opened workbook in place after a successful run. `--save-as` writes a copy after a successful run and must keep the same workbook extension as the opened workbook. `--save` and `--save-as` cannot be combined.

`run --direct` executes an argument-free, trace-disabled macro through `Excel.Run($MacroName)` without injecting the temporary harness module. It cannot be combined with `--arg` or `--trace`; those combinations fail before Excel starts. Direct runs return weaker VBA diagnostics because errors are surfaced by Excel COM rather than the xlflow harness. `run --fast` uses direct execution when the macro has no CLI arguments and trace is disabled; otherwise it falls back to the normal harness. `run --session` attaches to the workbook opened by `session start`.

`run --headless` is for AI agents, tests, and CI. Before Excel starts, xlflow scans the configured VBA source tree for GUI boundaries. If any boundary is found, the run fails with `gui_boundary_detected`, exit code `1`, and top-level `gui_boundaries` containing the detected file, line, kind, symbol, severity, message, and suggestion. `run --interactive` is for human-assisted Excel workflows. It runs with Excel visible and alerts enabled so a person can complete dialogs, message boxes, or forms. `--headless` and `--interactive` cannot be combined. `--timeout` defaults to `5m`; if a run exceeds the timeout, xlflow returns `macro_timeout` with exit code `1` and guidance that a dialog, form, file picker, or loop may still be waiting. Running without either mode keeps the legacy behavior except for the timeout.

`run` adds a `macro` object with `name`, `args`, and `duration_ms`. Failed macro runs return `macro_failed` with `error.source`, `error.number`, `error.message`, `error.line` when VBA exposes a non-zero `Erl` value, and `error.phase` when the failed phase is known. Stable run phases are `open_workbook`, `prepare_vbide`, `inject_harness`, `invoke_macro`, `save_result`, and `read_trace`. When Excel exposes enough information to distinguish a missing or invalid target macro from user-code failure, `run` returns `macro_not_found` instead of `macro_failed`. Plain-text success output must include the elapsed duration and whether the workbook was saved, copied, or left unchanged. Plain-text failure output must use the formatted message `Module line <n> Err <n>: <description>` when line and error number are available, and otherwise omit the `line <n>` segment. Because `run` injects a temporary VBA harness to measure duration while avoiding modal VBA runtime error dialogs, VBIDE access failures return an environment error such as `vbide_access_denied` and exit code `3`.

`run --trace` creates a fresh log under `.xlflow/traces`, calls `XlflowTrace.XlflowSetTraceFile` before the target macro, then reads trace events after execution. User VBA code writes events with `Call XlflowLog("message")`. If `XlflowTrace` is missing, `run --trace` temporarily injects it and reverts the helper before saving successful results. JSON output adds top-level `trace.enabled`, `trace.path`, `trace.events`, lifecycle metadata, and optional `trace.read_error`; each event has `timestamp`, `message`, and `raw`. Plain-text output prints trace events as `[timestamp] message` after the normal run logs. If the macro fails after writing trace events, those events are still returned. Trace read errors are reported in `trace.read_error` without changing the macro result. If a traced run fails with zero events, output should indicate that execution may have failed before reaching user trace calls.

`analyze` scans configured source directories without Excel COM for runtime-risk patterns. It returns top-level `analysis`; findings contain `code`, `severity`, `file`, `module`, `procedure`, `line`, `message`, `reason`, `suggestion`, and `nearby_code`. Findings are validation failures with exit code `1`.

`check` runs `lint`, `analyze`, then `doctor`. It continues after lint/analyze findings so source issues and environment status are returned together. JSON output includes top-level `check`, `issues`, `analysis`, and doctor diagnostics. Lint/analyze findings return exit code `1`; doctor/environment failure returns exit code `3`.

`macros` opens the configured workbook and discovers public runnable VBA entrypoints without executing user code. JSON output includes top-level `macros`, where each entry contains `module`, `name`, `qualified_name`, `kind` when available, and `args` when available. `macros --session` reads from the workbook opened by `session start`. Agents should use this command before guessing a `run` target.

`ui button add` opens the configured workbook and adds or updates an xlflow-managed Excel form-control button. The target worksheet is selected by `--sheet`; if it does not exist, the command fails with `sheet_not_found` unless `--create-sheet` is set. `--cell` is the top-left placement anchor, `--text` becomes the button caption, and `--macro` is assigned to the button `OnAction`. `--width` and `--height` are in Excel points and default to `160` and `40`. The stable internal button name is `xlflow.button.<id>`, where `<id>` is the normalized `--id` value or, when omitted, a normalized value derived from `--macro`. Re-running `add` with the same id updates the existing button instead of creating duplicates. `--verify-macro` checks the workbook VBIDE project for the macro before saving; missing macros fail with `macro_not_found`, and unavailable VBIDE access is an environment failure.

`ui button list` reports only xlflow-managed form-control buttons whose internal names start with `xlflow.button.`. When `--sheet` is provided, only that worksheet is inspected and a missing worksheet fails with `sheet_not_found`. `list` does not save the workbook.

`ui button remove` deletes an xlflow-managed form-control button by `--id`, optionally restricted to `--sheet`. Missing worksheets fail with `sheet_not_found`; missing buttons fail with `button_not_found`. `remove` saves the workbook only after a successful deletion.

`inspect-gui` scans configured source directories and reports GUI interaction boundaries without opening Excel. JSON output includes top-level `gui_boundaries`. Human output shows each boundary location, kind, symbol, and suggested refactor.

`attach --active` inspects the current active Excel workbook. It verifies that the active workbook path matches configured `excel.path` and reports top-level `workbook.path`, `workbook.configured_path`, `workbook.active`, and `workbook.matches_config`. In this version, `attach` does not change the connection target for `pull`, `push`, or `run`; it only validates the human-opened workbook.

`test` opens the configured workbook, discovers argument-free `Sub` procedures from the workbook VBIDE state, and runs procedures whose names start with `Test` or end with `_Test`. `--filter` uses exact procedure-name matching. `test --session` runs against the workbook opened by `session start`. Duplicate discovered test names, no discovered tests, missing filter targets, and VBA test failures are validation failures. Excel, COM, VBIDE, PowerShell, and script failures are environment failures.

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
- `analysis` for `analyze` and `check`
- `check` for `check`
- `run_diagnostic` for enriched `run` failures
- `trace` for traced `run`
- `session` for session status metadata
- `runner` for persistent runner module status
- `gui_boundaries` for `inspect-gui`, `run --headless` preflight failures, and `doctor` source summaries
- `ui` for `ui button` commands

`test` result objects contain `name`, `module`, `status`, `duration_ms`, and an optional `error`.

`ui button add` and `ui button remove` return `ui.button` with `id`, `name`, `sheet`, `text`, `macro`, `cell`, `left`, `top`, `width`, `height`, and `updated`. `ui button list` returns `ui.buttons` with the same fields for each managed button.

`diff` result objects contain `summary`, `sheets`, `cells`, and `vba`. Cell diffs contain `sheet`, `address`, `kind`, `before`, and `after`, where `kind` is `value` or `formula`. VBA diffs contain `file`, `kind`, and optional changed line details.

## Exit Codes

- `0`: success
- `1`: user-code or validation failure, including lint findings, analysis findings, GUI boundary preflight failures, macro failure, macro timeout, missing macro target, trace source removal refusal, missing UI sheets or buttons, VBA test failure, no tests found, missing filter targets, active workbook mismatches, and duplicate test names
- `2`: CLI argument or configuration error, including invalid `push`, `run`, `session`, `save`, and `runner` option combinations
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
- `VB007`: automation-hostile GUI boundaries such as file pickers, modal dialogs, UserForms, message pumps, or external process launches. JSON findings may include `kind`, `symbol`, and `suggestion`.

## Analysis Rules

- `VBA101`: object variable assignment likely missing `Set`
- `VBA102`: object-returning function assignment likely missing `Set`
- `VBA103`: object-returning function body likely missing `Set <FunctionName> = ...`
