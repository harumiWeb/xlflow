# xlflow Debugging Reference

This document describes the default debugging workflow for unclear VBA runtime and compile failures. Load it when `xlflow run` or `xlflow test` fails and the structured diagnostics alone do not identify the root cause.

## Default Order

Start with the stable path:

1. Run the macro normally with JSON output.
2. Read the returned error metadata, captured dialog text, and any diagnostic location.
3. Only add diagnostic instrumentation when the default diagnostics are not enough.

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
- `error.message`
- `error.location` when available
- `debug.events` when `XlflowDebug.Log` already exists in the code path

If those fields already identify the failing statement or bad API call, fix the source directly and rerun.

## Runtime Error Location Strategy: Line Numbers + Erl + XlflowDebug.Log

For unclear VBA runtime errors, the preferred diagnostic pattern is:

1. Add line numbers.
2. Add or use an error handler.
3. Log `Err.Number`, `Err.Description`, and especially `Erl` with `XlflowDebug.Log`.
4. Use the reported `Erl` value to locate the failing numbered statement.
5. Add targeted state logs only if the line is known but the root cause is still unclear.

Do not treat line numbers and `XlflowDebug.Log` as separate debugging tools. For runtime errors, line numbers are mainly useful because they make `Erl` meaningful.

`Erl` returns the last executed VBA line number before the runtime error. Without line numbers, `Erl` usually returns `0`, which is not useful for locating the failure.

## When To Add Line Numbers

Use line numbers when the default diagnostics do not identify the failing runtime statement.

Recommended sequence:

```bash
xlflow fmt --line-numbers add --write
xlflow push --fast --session --no-save --json
xlflow run <QualifiedMacro> --headless --session --json
```

Rules:

- Inspect the diff after adding line numbers, or review the git diff immediately after writing.
- Treat added line numbers as diagnostic instrumentation unless the project intentionally keeps them.
- Do not remove pre-existing line numbers that look intentional unless the user asks for that change.
- If `fmt --line-numbers remove` or `renumber` reports ambiguity around numeric labels, stop and report it instead of forcing a rewrite.

## Required Runtime Error Handler Pattern

When a runtime error location is unclear, add a temporary error handler that logs `Erl`.

Use this pattern:

```vb
Public Sub GenerateReport()
10  On Error GoTo ErrHandler

20  Dim total As Double
30  total = 1 / 0

40  Exit Sub

ErrHandler:
50  XlflowDebug.Log "Err.Number=" & CStr(Err.Number)
60  XlflowDebug.Log "Err.Description=" & Err.Description
70  XlflowDebug.Log "Erl=" & CStr(Erl)
80  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
```

After running this with xlflow, inspect `debug.events` or stderr.

Example result:

```text
Err.Number=11
Err.Description=Division by zero
Erl=30
```

Then map `Erl=30` back to the numbered source:

```vb
30  total = 1 / 0
```

This is the failing statement.

## Existing Error Handlers

If the procedure already has an error handler, do not blindly replace it.

Instead, add `XlflowDebug.Log` lines inside the existing handler before any `Resume`, `Resume Next`, `Exit Sub`, `Exit Function`, or custom error handling logic.

Example:

```vb
Public Sub GenerateReport()
10  On Error GoTo ErrHandler

20  Call BuildReport
30  Exit Sub

ErrHandler:
40  XlflowDebug.Log "GenerateReport failed"
50  XlflowDebug.Log "Err.Number=" & CStr(Err.Number)
60  XlflowDebug.Log "Err.Description=" & Err.Description
70  XlflowDebug.Log "Erl=" & CStr(Erl)

80  MsgBox Err.Description
End Sub
```

If the existing handler uses `Resume Next`, preserve that behavior unless the task explicitly requires changing error handling semantics.

Bad:

```vb
ErrHandler:
    XlflowDebug.Log "Erl=" & CStr(Erl)
    Err.Raise Err.Number, Err.Source, Err.Description
```

Do not add `Err.Raise` to an existing handler if doing so changes the macro's intended control flow.

Good:

```vb
ErrHandler:
    XlflowDebug.Log "Err.Number=" & CStr(Err.Number)
    XlflowDebug.Log "Err.Description=" & Err.Description
    XlflowDebug.Log "Erl=" & CStr(Erl)
    Resume Next
```

## `XlflowDebug.Log`

Use `XlflowDebug.Log` for two different purposes:

1. Error location reporting from handlers.
2. Targeted state logging around suspected code paths.

For runtime error location, always prefer this minimum set:

```vb
XlflowDebug.Log "Err.Number=" & CStr(Err.Number)
XlflowDebug.Log "Err.Description=" & Err.Description
XlflowDebug.Log "Erl=" & CStr(Erl)
```

For root-cause analysis after the failing line is known, add targeted logs such as:

```vb
XlflowDebug.Log "entered GenerateReport"
XlflowDebug.Log "rowCount=" & CStr(rowCount)
XlflowDebug.Log "targetPath=" & targetPath
XlflowDebug.Log "worksheet=" & ws.Name
XlflowDebug.Log "before SaveAs"
```

Use targeted logs for:

- procedure entry
- branch selection
- variable state
- loop progress
- file paths
- worksheet names
- workbook names
- code immediately before destructive workbook changes
- existing error handlers

`XlflowDebug.Log` is preferred over raw `Debug.Print` for agent workflows because `xlflow run` and `xlflow test` stream it to stderr and return recent lines in top-level `debug`.

## Recommended Runtime Debugging Loop

```bash
xlflow run <QualifiedMacro> --headless --session --json
```

If the runtime error location is still unclear:

```bash
xlflow fmt --line-numbers add --write
xlflow push --fast --session --no-save --json
```

Then add temporary `Erl` logging to the relevant error handler:

```vb
XlflowDebug.Log "Err.Number=" & CStr(Err.Number)
XlflowDebug.Log "Err.Description=" & Err.Description
XlflowDebug.Log "Erl=" & CStr(Erl)
```

Rerun:

```bash
xlflow run <QualifiedMacro> --headless --session --json
```

Then:

1. Read `debug.events` or stderr.
2. Find `Erl=<number>`.
3. Locate that numbered line in the source.
4. Fix the failing statement if the cause is obvious.
5. If the line is known but the cause is unclear, add targeted `XlflowDebug.Log` state logs around that line.
6. Push and rerun.
7. Remove temporary diagnostic logs before finalizing.

## Compile Errors

For compile errors, `Erl` is not useful because the code did not run.

Use the structured compile diagnostics first. If xlflow reports a VBE-selected source location, patch that location directly.

If compile diagnostics do not include a location, inspect the compile error message and recently changed modules. Line numbers are usually less useful for compile errors than for runtime errors.

## Cleanup

After the cause is fixed:

- remove temporary `XlflowDebug.Log` calls
- decide whether line numbers stay in the project
- if line numbers were diagnostic-only, remove them:

```bash
xlflow fmt --line-numbers remove --write
```

- if the project wants stable numbering, normalize it:

```bash
xlflow fmt --line-numbers renumber --write
```

Then push and rerun the final validation:

```bash
xlflow push --fast --session --no-save --json
xlflow run <QualifiedMacro> --headless --session --json
```

## Summary

Use this order:

1. Run normally and inspect structured diagnostics.
2. For unclear runtime errors, add line numbers.
3. Add or reuse an error handler that logs `Err.Number`, `Err.Description`, and `Erl` with `XlflowDebug.Log`.
4. Use `Erl` to identify the failing numbered statement.
5. Add targeted `XlflowDebug.Log` state logs only when the failing line is known but the cause is still unclear.
6. Fix the source.
7. Remove temporary instrumentation unless the project intentionally keeps it.
