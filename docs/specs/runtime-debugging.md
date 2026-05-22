# Runtime Debugging Hardening

## Scope

This spec defines the xlflow behavior that helps AI agents debug VBA runtime failures without relying on workbook-only state or implicit macro naming assumptions.

## Trace Lifecycle

`xlflow trace enable` is source-aware in configured projects. When the command uses `excel.path` from `xlflow.toml`, it injects or replaces the workbook module `XlflowTrace` and writes the same bundled module source to `<src.modules>/XlflowTrace.bas` as UTF-8 without BOM. `xlflow trace inject` remains a compatibility alias for `trace enable`.

This keeps `push` from deleting the trace module on the next source-to-workbook sync. The generated source file is owned by xlflow and may be replaced by a later `trace enable` run.

If `.xlflow/session.json` already points at the same workbook, trace lifecycle commands should reuse that live session workbook by default instead of opening a second hidden copy. This prevents false-success trace injection where the visible session workbook still lacks `XlflowTrace`.

When an explicit workbook argument is provided and xlflow cannot load project configuration, `trace enable <workbook>` may operate on the workbook only. That standalone mode exists for one-off workbook inspection and does not promise source persistence.

JSON output for configured project injection includes source metadata:

```json
{
  "source": {
    "path": "src/modules/XlflowTrace.bas",
    "updated": true
  }
}
```

`xlflow trace disable` removes the workbook helper and removes the generated source helper only when it still matches xlflow's bundled trace module. If the source helper has been modified, disable refuses with `trace_source_modified` unless `--force` is set. `xlflow trace status` reports workbook helper presence, source helper presence, whether source matches the bundled helper, and the trace log directory. `xlflow trace clean` removes `.xlflow/traces`.

## Run Failure Phases

`xlflow run` reports the phase that failed so callers can distinguish environment setup failures from user-code failures. Stable phase names are:

- `open_workbook`
- `prepare_vbide`
- `compile_vba`
- `verify_macro`
- `inject_harness`
- `invoke_macro`
- `save_result`
- `read_trace`

The phase is included in JSON error metadata. Plain-text output remains short, but failures should include enough context for a user or agent to decide whether to inspect configuration, VBIDE access, macro names, source code, or trace output.

When Excel exposes enough information to distinguish a missing or invalid macro target from user-code failure, xlflow reports a target-specific error code instead of generic `macro_failed`.

For `macro_failed` during `invoke_macro`, xlflow may add top-level `run_diagnostic`. Diagnostics include location, nearby source, trace context, likely cause, and suggestion when source analysis can match the failure to a known runtime-risk pattern.

## Runtime Modal Suppression

Default `xlflow run` suppresses VBA runtime error dialogs owned by the Excel process during macro invocation. xlflow closes the dialog, returns a structured CLI failure, and adds `run_diagnostic.kind = "runtime"` with dialog metadata and VBE selection context when Excel exposes it.

This applies to both the normal temporary-harness path and `run --direct`. Interactive mode does not implicitly opt out of runtime modal suppression; human/debug workflows must opt out explicitly with `--gui-compile-errors` if they want VBA error dialogs to stay in the GUI.

When a suppressed runtime dialog still yields structured VBA error data, xlflow keeps the existing failure codes such as `macro_failed`, `macro_not_found`, and `macro_disabled`, with `error.phase = "invoke_macro"`. If the dialog is the only reliable signal, xlflow may populate `error.message` from the localized dialog text and infer `error.number` from the dialog when possible.

## Diagnostic Compile Mode

`xlflow run --diagnostic` adds a VBE compile step before macro verification and invocation. It is intended for agent debugging when source preflight did not catch a compile-time issue but Excel would otherwise surface a modal VBE dialog.

Diagnostic mode starts a Win32 watcher for top-level windows owned by the Excel process, executes VBE Compile through the VBE command bars, and closes the compile dialog after collecting its child control text. Dialog text is returned as localized opaque text; xlflow does not parse or translate Japanese or English compile messages.

Compile failures return `vba_compile_failed` with `error.phase = "compile_vba"` and validation exit code `1`. `run_diagnostic.kind = "compile"` includes the dialog message, VBE selection location, nearby code, and dialog metadata when available. `--diagnostic --direct` is invalid. `--diagnostic --fast` remains valid but disables the direct fast path so the run can keep structured diagnostics. `--gui-compile-errors` disables this compile watcher and the default runtime modal suppression path so VBA errors remain in the GUI intentionally.

## Runtime Source Analysis

`xlflow analyze` scans configured source directories without Excel COM. The v1 analyzer is deliberately pattern-based and detects likely missing `Set` assignments for object variables and object-returning functions, statically-known object/member mismatches such as `Worksheet.DisplayGridlines`, and missing xlflow trace-helper dependencies when source-controlled VBA calls `XlflowLog` or `XlflowSetTraceFile` without a Public standard-module definition. Stable analyzer codes are `VBA101`, `VBA102`, `VBA103`, `VBA104`, `VBA105`, and `VBA106`.

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

`xlflow inspect-gui` exposes the same boundary report without running Excel. `xlflow attach --active` verifies that the human-opened active workbook matches the configured workbook before an interactive workflow continues.

When `run` or `test` invokes user VBA, xlflow also injects a workbook-scoped runtime mode marker before execution and restores the reserved names afterward. New scaffolded projects expose that state through `XlflowRuntime.bas`, so VBA can branch with helpers such as `XlflowRuntime.ModeName()`, `XlflowRuntime.IsHeadless()`, `XlflowRuntime.IsAgent()`, and `XlflowRuntime.IsTest()`. The workbook-scoped marker is the primary contract; `Environ$("XLFLOW_MODE")` is only a secondary fallback for manual helper adoption in older projects or wrapper-driven runs.

New scaffolded projects also include `XlflowDebug.bas`. `XlflowDebug.Log` is the explicit workbook-side logging surface for terminal-visible debug lines. During xlflow `run` and `test`, the runtime injection layer writes a temporary `__XLFLOW_DEBUG_PIPE__` workbook defined name before user VBA starts and restores the prior state afterward. `XlflowDebug.Log` always writes to the native VBA Immediate Window, and when that debug pipe marker is present it also emits newline-delimited JSON events to xlflow's debug stream so stderr and final JSON output can include the same log lines without corrupting stdout.

For runtime debugging sessions that need more than one workbook-backed command, keep the workbook attached through `session start` and reuse that session for `push --fast --session --no-save`, `run --session`, `test --session`, inspect, and save steps. Reopening the workbook separately for each step is intentionally not the preferred path because it slows down debugging and makes the live workbook state harder to reason about. Use a fresh non-session open only when the reopen boundary itself is the thing you are investigating.

## Empty Trace Guidance

`xlflow run --trace` returns all trace events written before failure. Trace logs are written under `.xlflow/traces`. If the workbook does not already contain `XlflowTrace`, xlflow may inject it temporarily and revert it before saving successful results. Source preflight for configured `run --trace` may ignore missing-helper findings that are satisfied by this temporary injection path, while `push` and non-trace configured runs keep treating those findings as blocking. If a traced run fails with zero events, output indicates that execution may have failed before reaching user trace calls.

The bundled AI agent skill instructs agents to add trace logs at procedure entry, important branches, external file access, destructive operations, and error handlers.

## PowerShell Host Diagnostics

Excel COM-backed commands return top-level `bridge` metadata with `host`, `edition`, and `version`. This identifies the xlflow PowerShell bridge host only.

When workbook VBA launches its own external PowerShell process, agents should not assume that it matches `bridge.host`. They should inspect the VBA command string or log the resolved executable from workbook code.

Windows review checklist:

1. Check `bridge.host` to confirm which PowerShell xlflow itself used.
2. Check the workbook-side command or resolved executable if VBA launches `powershell.exe`, `pwsh.exe`, or another shell.
3. Prefer one host consistently when debugging encoding or environment differences across external-process VBA flows.

## Agent Progress Output

Excel COM-backed commands always report in-flight progress on stderr. Interactive terminals show a spinner, and `--json` or non-interactive runs keep that same progress channel on stderr so stdout stays reserved for the final human output or JSON envelope.

Agents should use normal commands such as `xlflow pull --json`, `xlflow push --json`, and `xlflow run --json`, ignore transient spinner frames on stderr, and synchronize on process exit instead of any interim progress text.

## Bundled Skill Workflow Guidance

The bundled AI agent skill must make xlflow's source-first workflow explicit. In configured projects, agents should treat the configured source directories as authoritative unless the user says the workbook has newer VBA or the source tree is missing or stale. In those cases, agents should run `xlflow pull --json` before editing and then continue from source files.

The skill must tell agents to use `xlflow macros --json` and a discovered `qualified_name` before running a macro when the entrypoint is unclear. Agents should not assume default names such as `Main.Run` unless discovery, tests, docs, or prior command output prove that entrypoint.

The skill must distinguish environment/setup failures from user-code failures. For setup phases such as `open_workbook`, `prepare_vbide`, and `inject_harness`, agents should run `xlflow doctor --json` before changing VBA source. For `invoke_macro` failures, agents should inspect VBA error metadata and trace events before patching source.
