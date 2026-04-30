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

`xlflow lint` detects interactive VBA patterns that block unattended CLI execution:

- `Application.GetOpenFilename`
- `Application.FileDialog`
- `InputBox`
- modal `MsgBox`

Findings explain that xlflow-oriented macros should prefer explicit `run --arg` values, environment variables, configuration cells, or deterministic paths over UI prompts.

## Empty Trace Guidance

`xlflow run --trace` returns all trace events written before failure. If a traced run fails with zero events, output indicates that execution may have failed before reaching user trace calls.

The bundled AI agent skill instructs agents to add trace logs at procedure entry, important branches, external file access, destructive operations, and error handlers.
