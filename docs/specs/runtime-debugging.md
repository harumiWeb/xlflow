# Runtime Debugging Hardening

## Scope

This spec defines the xlflow behavior that helps AI agents debug VBA runtime failures without relying on workbook-only state or implicit macro naming assumptions.

## Trace Injection Persistence

`xlflow trace inject` is source-aware in configured projects. When the command uses `excel.path` from `xlflow.toml`, it injects or replaces the workbook module `XlflowTrace` and writes the same bundled module source to `<src.modules>/XlflowTrace.bas` as UTF-8 without BOM.

This keeps `push` from deleting the trace module on the next source-to-workbook sync. The generated source file is owned by xlflow and may be replaced by a later `trace inject` run.

When an explicit workbook argument is provided and xlflow cannot load project configuration, `trace inject <workbook>` may operate on the workbook only. That standalone mode exists for one-off workbook inspection and does not promise source persistence.

JSON output for configured project injection includes source metadata:

```json
{
  "source": {
    "path": "src/modules/XlflowTrace.bas",
    "updated": true
  }
}
```

## Run Failure Phases

`xlflow run` reports the phase that failed so callers can distinguish environment setup failures from user-code failures. Stable phase names are:

- `open_workbook`
- `prepare_vbide`
- `inject_harness`
- `invoke_macro`
- `save_result`
- `read_trace`

The phase is included in JSON error metadata. Plain-text output remains short, but failures should include enough context for a user or agent to decide whether to inspect configuration, VBIDE access, macro names, source code, or trace output.

When Excel exposes enough information to distinguish a missing or invalid macro target from user-code failure, xlflow reports a target-specific error code instead of generic `macro_failed`.

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

## Empty Trace Guidance

`xlflow run --trace` returns all trace events written before failure. If a traced run fails with zero events, output indicates that execution may have failed before reaching user trace calls.

The bundled AI agent skill instructs agents to add trace logs at procedure entry, important branches, external file access, destructive operations, and error handlers.

## Agent Keepalive Output

`--keepalive` on Excel COM-backed commands is for AI agents and task runners that may stop waiting when Excel COM operations are silent for too long. It is available on `new`, `doctor`, `attach`, `pull`, `push`, `trace inject`, `run`, `macros`, `ui button add/list/remove`, and `test`. Keepalive writes only to stderr. Stdout remains reserved for normal human output or the JSON envelope when `--json` is set.

Heartbeat output starts immediately with `xlflow: <command> still running... elapsed=0s` and repeats at `--keepalive-interval`, which defaults to `5s`. At command completion, xlflow writes `XLFLOW_DONE status=success command=<command>` or `XLFLOW_DONE status=failed command=<command> code=<error-code>` when a structured error code is available.

Agents should use `--keepalive --json` for long Excel COM-backed calls, wait for process exit, and treat the `XLFLOW_DONE` marker as the synchronization point before starting the next workbook-dependent command.

## Bundled Skill Workflow Guidance

The bundled AI agent skill must make xlflow's source-first workflow explicit. In configured projects, agents should treat the configured source directories as authoritative unless the user says the workbook has newer VBA or the source tree is missing or stale. In those cases, agents should run `xlflow pull --json` before editing and then continue from source files.

The skill must tell agents to use `xlflow macros --json` and a discovered `qualified_name` before running a macro when the entrypoint is unclear. Agents should not assume default names such as `Main.Run` unless discovery, tests, docs, or prior command output prove that entrypoint.

The skill must distinguish environment/setup failures from user-code failures. For setup phases such as `open_workbook`, `prepare_vbide`, and `inject_harness`, agents should run `xlflow doctor --json` before changing VBA source. For `invoke_macro` failures, agents should inspect VBA error metadata and trace events before patching source.
