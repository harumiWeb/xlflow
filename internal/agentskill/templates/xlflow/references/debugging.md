# xlflow Debugging Reference

This document describes the default debugging workflow for unclear VBA runtime and compile failures. Load it when `xlflow run` or `xlflow test` fails and the structured diagnostics alone do not identify the root cause.

## Default Order

Start with the stable path:

1. Run the macro normally with JSON output.
2. Read the returned error metadata and any captured dialog/diagnostic location.
3. Only add line numbers and extra logging when the default diagnostics are not enough.

Do not switch to VBE `Debug` mode as the first step. xlflow's default `run` behavior is intentionally biased toward stable unattended execution and structured failure reporting.

## First Pass

Prefer the normal reproduction first:

```bash
xlflow run --headless --session --json
xlflow run <QualifiedMacro> --headless --session --json
```

If the project is not using headless mode for that path, use the same command shape the task already depends on. The important point is to inspect the returned diagnostics before changing source.

Read:

- `error.code`
- `error.phase`
- VBA error number, description, and line when present
- `debug.events` when `XlflowDebug.Log` already exists in the code path

If those fields already identify the failing statement or bad API call, fix the source directly and rerun. Do not add extra instrumentation just because a command failed.

## When To Add Line Numbers

Use line numbers only when the default diagnostics still leave the failing location unclear.

Recommended sequence:

```bash
xlflow fmt --line-numbers add --write
xlflow push --fast --session --no-save --json
xlflow run <QualifiedMacro> --headless --session --json
```

Rules:

- Inspect the diff after `fmt --line-numbers add` before `--write`, or review the git diff immediately after writing.
- Treat added line numbers as diagnostic instrumentation unless the project intentionally keeps them.
- Do not remove pre-existing line numbers that look intentional unless the user asks for that change.
- If `fmt --line-numbers remove` or `renumber` reports ambiguity around numeric labels, stop and report it instead of forcing a rewrite.

## `XlflowDebug.Log`

When line numbers still are not enough, add targeted `XlflowDebug.Log` calls around the suspected path.

Use it for:

- procedure entry
- branch selection
- variable state
- loop progress
- file paths or worksheet names
- code immediately before destructive workbook changes
- error handlers

Examples:

```vb
XlflowDebug.Log "entered GenerateReport"
XlflowDebug.Log "rowCount=" & CStr(rowCount)
XlflowDebug.Log "targetPath=" & targetPath
XlflowDebug.Log "before SaveAs"
```

When the failure is runtime-error-driven, log the VBA error state from the handler:

```vb
Public Sub GenerateReport()
    On Error GoTo ErrHandler

    Dim total As Double
    total = 1 / 0
    Exit Sub

ErrHandler:
    XlflowDebug.Log "Err.Number=" & CStr(Err.Number)
    XlflowDebug.Log "Err.Description=" & Err.Description
    XlflowDebug.Log "Erl=" & CStr(Erl)
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
```

`XlflowDebug.Log` is preferred over raw `Debug.Print` for agent workflows because `run` and `test` stream it to stderr and return the recent lines in top-level `debug`.

## Recommended Loop

```bash
xlflow run <QualifiedMacro> --headless --session --json
xlflow fmt --line-numbers add --write
xlflow push --fast --session --no-save --json
xlflow run <QualifiedMacro> --headless --session --json
```

Then:

1. Use the numbered source and `Erl` to find the failing statement.
2. Add targeted `XlflowDebug.Log` calls if the root cause is still unclear.
3. Push and rerun.
4. Fix the source.
5. Remove temporary `XlflowDebug.Log` calls before finalizing.

## Cleanup

After the cause is fixed:

- remove temporary `XlflowDebug.Log` calls
- decide whether line numbers stay in the project
- if they were diagnostic-only, remove them:

```bash
xlflow fmt --line-numbers remove --write
```

- if the project wants stable numbering, normalize it:

```bash
xlflow fmt --line-numbers renumber --write
```

## Summary

Use this order:

1. `run`
2. inspect structured diagnostics
3. `fmt --line-numbers add --write` only when location is still unclear
4. targeted `XlflowDebug.Log`
5. rerun, fix, and remove temporary instrumentation
