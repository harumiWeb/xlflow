# Runtime Debugging Hardening

## Scope

This spec defines the xlflow behavior that helps AI agents debug VBA runtime failures without relying on workbook-only state or implicit macro naming assumptions.

## Recommended Workflow

The legacy debugging command and removed run flag are gone. The supported workbook-side debugging surface is `XlflowDebug.Log`, and the supported machine-readable execution surface is `xlflow run --json`.

Use `XlflowDebug.Log` inside VBA for execution-state visibility:

```vb
XlflowDebug.Log "message"
XlflowDebug.Log "row count", rowCount
XlflowDebug.Log "current sheet", ws.Name
```

During `run` and `test`, xlflow injects the temporary debug pipe marker used by `XlflowDebug.Log` and returns those lines through stderr streaming and the final JSON envelope.

## Run JSON Verbosity

`xlflow run --json` keeps diagnostic behavior enabled by default but uses a compact failure payload optimized for AI-agent and normal development loops. The default run JSON keeps the high-signal fields needed to answer what failed, where it failed, what workbook/session state matters next, and what action is suggested next.

For failed macro runs, the default JSON contract keeps top-level `status`, `command`, `error`, `macro.name`, `macro.duration_ms`, `location`, `session`, `target`, `warnings`, and `suggestion`. The failure `location` is promoted to a top-level field so callers do not need to parse xlflow's internal diagnostic namespace just to find the relevant source file and line.

`xlflow run --json --verbose` preserves the broader diagnostic payload used for xlflow internals, dialog-watcher debugging, and bug reports. `--verbose` controls output verbosity only; it does not change compile/runtime diagnostic collection, modal dialog suppression, or location capture behavior.

## Run Failure Phases

`xlflow run` reports the phase that failed so callers can distinguish environment setup failures from user-code failures. Stable phase names are:

- `open_workbook`
- `prepare_vbide`
- `compile_vba`
- `verify_macro`
- `inject_harness`
- `invoke_macro`
- `save_result`

The phase is included in JSON error metadata. Plain-text output remains short, but failures should include enough context for a user or agent to decide whether to inspect configuration, VBIDE access, macro names, source code, or debug output.

When Excel exposes enough information to distinguish a missing or invalid macro target from user-code failure, xlflow reports a target-specific error code instead of generic `macro_failed`.

For `macro_failed` during `invoke_macro`, xlflow may add top-level `run_diagnostic`. Diagnostics include location, nearby source, debug context, likely cause, and suggestion when source analysis can match the failure to a known runtime-risk pattern. In the default JSON contract, `suggestion` and the best available `location` are promoted to top-level fields; `run_diagnostic` remains a verbose/debug-oriented namespace.

## Runtime Modal Suppression

Default `xlflow run` suppresses VBA runtime error dialogs owned by the Excel process during macro invocation. xlflow closes the dialog, returns a structured CLI failure, and adds `run_diagnostic.kind = "runtime"` with dialog metadata and VBE selection context when Excel exposes it.

For `.NET` bridge runs, VBE selection capture is best-effort and timeout-bounded. When Excel exposes the active VBE code pane, `run_diagnostic.location` may include `confidence`, `method`, `source_path`, `component`, `component_type`, `procedure`, `line`, `column`, `end_line`, `end_column`, and `text`. Field names follow xlflow's `snake_case` JSON convention. Verified line values are source-file line numbers, adjusted for exported metadata that VBE hides such as `Attribute VB_*`; column values are omitted when VBE only exposes an unreliable whole-line selection. Capture failures do not change the command failure code; they are reported under `run_diagnostic.location_capture.attempts` with timings such as `before_dialog_action` or `after_dialog_action` when verbose output is requested.

Default `run --json` does not emit full dialog/window/control snapshots, duplicate `dialog` and `dialogs` payloads, low-level dialog-watcher metadata, workbook/bridge/runtime debug metadata, or location-capture attempt details. Verbose run JSON keeps those details for troubleshooting and bug reports. When verbose dialog snapshots are emitted, `dialogs` is the preferred full snapshot field; callers should not depend on duplicated single-dialog payloads.

The .NET bridge does not require a runtime dialog to be visible. Excel can defer
painting an owned dialog until the Excel window receives focus. Unattended
suppression prefers the explicit End action because Debug can leave VBE in break
mode and prevent later COM attachment.

This applies to both the normal temporary-harness path and `run --direct`. Interactive mode does not implicitly opt out of runtime modal suppression; human/debug workflows must opt out explicitly with `--gui-compile-errors` if they want VBA error dialogs to stay in the GUI.

When a suppressed runtime dialog still yields structured VBA error data, xlflow keeps the existing failure codes such as `macro_failed`, `macro_not_found`, and `macro_disabled`, with `error.phase = "invoke_macro"`. If the dialog is the only reliable signal, xlflow may populate `error.message` from the localized dialog text and infer `error.number` from the dialog when possible.

Timeout is intentionally weaker than runtime dialog suppression. xlflow returns
`macro_timeout` with a valid JSON envelope and actionable suggestions, but it
does not attempt synchronous COM cleanup while Excel is still busy. Timeout
diagnostics therefore imply `vba_may_still_be_running`: the workbook and any
attached session should be treated as dirty until Excel is reset or the workbook
is reopened.

## Diagnostic Compile Mode

`xlflow run --diagnostic` adds a VBE compile step before macro verification and invocation. It is intended for agent debugging when source preflight did not catch a compile-time issue but Excel would otherwise surface a modal VBE dialog.

Diagnostic mode starts a Win32 watcher for top-level windows owned by the Excel process, executes VBE Compile through the VBE command bars, and closes the compile dialog after collecting its child control text. Dialog text is returned as localized opaque text; xlflow does not parse or translate Japanese or English compile messages.

Compile failures return `vba_compile_failed` with `error.phase = "compile_vba"` and validation exit code `1`. `run_diagnostic.kind = "compile"` includes the dialog message, VBE selection location, nearby code, and dialog metadata when available. Default `run --json` still promotes the best available `location` to the top level and keeps the rest of the compile diagnostic detail behind `--verbose`. `push` compile failures use the same `.NET` selection capture under `push_diagnostic.kind = "compile"` after source import and before the workbook is saved. In the `.NET` bridge, the selection capture runs before the dialog is dismissed and retries once immediately after dismissal only if no meaningful location was captured. `--diagnostic --direct` is invalid. `--diagnostic --fast` remains valid but disables the direct fast path so the run can keep structured diagnostics. `--gui-compile-errors` disables this compile watcher and the default runtime modal suppression path so VBA errors remain in the GUI intentionally.

## Runtime Source Analysis

`xlflow analyze` scans configured source directories without Excel COM. The analyzer uses `tree-sitter-vba` to build per-file and per-procedure context from declarations, parameters, function and property returns, labels, assignments, calls, and member access. It detects likely missing `Set` assignments for object variables and object-returning functions, statically-known Excel object/member mismatches, removed legacy helper APIs such as `XlflowLog` and `XlflowSetTraceFile`, `Range.Find` results used without `Nothing` guards, object variables used before obvious initialization, Application state leaks, error-handler fallthrough, unqualified Excel object access, and selected opt-in semantic risks. Stable analyzer codes are `VBA101` through `VBA106` and `VBA201` through `VBA211`.

`xlflow check` aggregates `lint`, `analyze`, and `doctor`. It continues after lint/analyze findings and returns all cheap source feedback before reporting Excel COM doctor status.

## Macro Entrypoint Discovery

xlflow provides a non-executing macro discovery command. The command reads the configured workbook and returns runnable public entrypoints in machine-readable form.

Each discovered entrypoint includes:

- module name
- procedure name
- fully qualified macro name
- procedure kind when available
- argument count or argument signature when available

Agents should use this discovery result before guessing a `run` target.

## Automation-Hostile VBA Patterns

xlflow treats GUI operations as explicit boundaries instead of trying to automate them invisibly. The same source scanner is used by `lint`, `doctor`, `inspect-gui`, and `run --headless`.

- `Application.GetOpenFilename`
- `Application.GetSaveAsFilename`
- `Application.FileDialog`
- `InputBox`
- modal `MsgBox`
- `UserForm.Show` and modal `.Show vbModal`
- `DoEvents`
- `Shell`
- `CreateObject("WScript.Shell").Popup`

Each boundary reports `file`, `line`, `kind`, `symbol`, `severity`, `message`, and `suggestion`. Stable `kind` values are `file_picker`, `modal_dialog`, `user_form`, `external_process`, and `message_pump`.

Findings explain that xlflow-oriented macros should prefer explicit `run --arg` values, environment variables, configuration cells, or deterministic paths over UI prompts. GUI entrypoints may remain available for humans, but the core business logic should be extractable into parameterized procedures that can run headlessly.

## Headless and Interactive Run Modes

`xlflow run --headless` is the default recommendation for AI agents and CI. It scans source before starting Excel. If GUI boundaries are present, it fails with `gui_boundary_detected` and returns top-level `gui_boundaries` so the agent can explain why execution was refused.

`xlflow run --interactive` is for human-assisted sessions. It runs Excel visibly with alerts enabled, allowing a person to complete file pickers, message boxes, or UserForms. `--timeout` defaults to five minutes; timeout failures return `macro_timeout` and should be interpreted as a possible unresolved dialog, form, file picker, or long-running loop.

`xlflow inspect-gui` exposes the same boundary report without running Excel. `xlflow session attach` adopts the human-opened configured workbook as the live xlflow session before an interactive workflow continues. The legacy `xlflow attach --active` command is deprecated and only validates the active workbook.

When `run` or `test` invokes user VBA, xlflow also injects a workbook-scoped runtime mode marker before execution and restores the reserved names afterward. New scaffolded projects expose that state through `XlflowRuntime.bas`, so VBA can branch with helpers such as `XlflowRuntime.ModeName()`, `XlflowRuntime.IsHeadless()`, `XlflowRuntime.IsAgent()`, and `XlflowRuntime.IsTest()`. The workbook-scoped marker is the primary contract; `Environ$("XLFLOW_MODE")` is only a secondary fallback for manual helper adoption in older projects or wrapper-driven runs.

New scaffolded projects also include `XlflowDebug.bas`. `XlflowDebug.Log` is the explicit workbook-side logging surface for terminal-visible debug lines. During xlflow `run` and `test`, the runtime injection layer writes a temporary `__XLFLOW_DEBUG_PIPE__` workbook defined name before user VBA starts and restores the prior state afterward. `XlflowDebug.Log` always writes to the native VBA Immediate Window, and when that debug pipe marker is present it also emits newline-delimited JSON events to xlflow's debug stream so stderr and final JSON output can include the same log lines without corrupting stdout.

For runtime debugging sessions that need more than one workbook-backed command, keep the workbook attached through `session start` or `session attach` and reuse that session for `push --fast --session --no-save`, `run --session`, `test --session`, inspect, and save steps. Reopening the workbook separately for each step is intentionally not the preferred path because it slows down debugging and makes the live workbook state harder to reason about. Use a fresh non-session open only when the reopen boundary itself is the thing you are investigating.

## Debug Logging Guidance

`xlflow run --json` returns debug events emitted through `XlflowDebug.Log`. If a failed run has insufficient context, add `XlflowDebug.Log` calls at procedure entry, important branches, external file access, destructive operations, and error handlers, then rerun with `--json`.

The bundled AI agent skill instructs agents to prefer `XlflowDebug.Log` plus structured `run` output for runtime debugging.

## External PowerShell Host Diagnostics

Excel COM-backed commands return top-level `.NET` bridge metadata such as `name`, `version`, `protocol_version`, `runtime`, and `architecture`.

When workbook VBA launches its own external PowerShell process, agents should not assume that it matches xlflow's bridge process. They should inspect the VBA command string or log the resolved executable from workbook code.

Windows review checklist:

1. Check top-level `bridge` metadata to confirm which xlflow bridge provider handled the command.
2. Check the workbook-side command or resolved executable if VBA launches `powershell.exe`, `pwsh.exe`, or another shell.
3. Prefer one host consistently when debugging encoding or environment differences across external-process VBA flows.

## Agent Progress Output

Excel COM-backed commands report progress on stderr. Interactive stderr terminals show a spinner; non-interactive or `--json` runs fall back to a single line of progress on stderr so stdout stays reserved for the final human output or JSON envelope. Commands that stream UI or debug events to stderr may suppress separate progress output so those event lines remain parseable.

Agents should use normal commands such as `xlflow pull --json`, `xlflow push --json`, and `xlflow run --json`, treat stderr progress as advisory only, and synchronize on process exit instead of any interim progress text.

## Bundled Skill Workflow Guidance

The bundled AI agent skill must make xlflow's source-first workflow explicit. In configured projects, agents should treat the configured source directories as authoritative unless the user says the workbook has newer VBA or the source tree is missing or stale. In those cases, agents should run `xlflow pull --json` before editing and then continue from source files.

The skill must tell agents to use `xlflow macros --json` and a discovered `qualified_name` before running a macro when the entrypoint is unclear. Agents should not assume default names such as `Main.Run` unless discovery, tests, docs, or prior command output prove that entrypoint.

The skill must distinguish environment/setup failures from user-code failures. For setup phases such as `open_workbook`, `prepare_vbide`, and `inject_harness`, agents should run `xlflow doctor --json` before changing VBA source. For `invoke_macro` failures, agents should inspect VBA error metadata and debug events before patching source.
