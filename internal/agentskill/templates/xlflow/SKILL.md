---
name: xlflow
description: Use when Codex or another AI agent needs to edit, test, debug, or validate Excel VBA workbooks with xlflow. Provides the safe VBA development workflow for xlflow projects, including pull/push, lint, run, trace, test, diff, failure handling, and final reporting rules.
---

# xlflow Skill

## Purpose

Use xlflow as the proof loop for Excel VBA work. Do not treat generated VBA as complete until the workbook has been exported or inspected, source has been changed, changes have been imported, and the relevant macro or tests have been run.

## Standard Workflow

1. Inspect the project.
   - Read `xlflow.toml`.
   - Treat the configured source directories as authoritative when `xlflow.toml` exists and the user has not said the workbook contains newer VBA.
   - Run `xlflow doctor --json` when Excel, COM, VBIDE, or macro execution looks suspicious.
   - Run `xlflow pull --json` before editing when the workbook is the current source of truth.

2. Edit source files.
   - Prefer `.bas`, `.cls`, and `.frm` files under the configured source directories.
   - Do not edit binary workbooks directly unless the task is explicitly workbook-state only.

3. Import and check.
   - Run `xlflow push --keepalive --json` after source edits.
   - Run `xlflow lint --json` and fix reported issues before finalizing.

4. Execute behavior.
   - Prefer `xlflow test --json`.
   - Use `xlflow test --filter <name> --json` while iterating on one failing test.
   - If the macro entrypoint is unclear, run `xlflow macros --json` before choosing a target.
   - If no tests exist, run the target macro with `xlflow run <MacroName> --keepalive --json`.
   - Prefer `xlflow run <MacroName> --headless --keepalive --json` for unattended agent work.
   - Use `xlflow run <MacroName> --interactive --json` only when a human can operate Excel dialogs or forms.
   - Use `xlflow run <MacroName> --trace --json` when debugging runtime behavior or workbook mutation.

5. Compare results.
   - Use `xlflow diff <before> <after> --json` for workbook state changes.
   - Add `--vba-before <dir> --vba-after <dir>` when exported source changes also need review.

6. Repeat until the command results prove the task.

## Project Orientation

Before editing, decide what is authoritative:

- If `xlflow.toml` exists and source files are present, edit the configured source tree and use `xlflow push --json` to update the workbook.
- If the user says the workbook has the latest VBA, or source files are missing or stale, run `xlflow pull --json` first and then edit source files.
- Do not mix direct workbook edits with source edits in the same task unless the requested change is workbook-state only and no VBA source change is needed.
- After `xlflow trace inject --json`, remember that `XlflowTrace.bas` is generated xlflow support code. Do not rewrite it by hand unless the user is changing xlflow itself.

Before running a macro, decide the runnable entrypoint:

- Run `xlflow macros --json` when the macro name is not already proven by tests, docs, or prior command output.
- Use a listed `qualified_name` from `xlflow macros --json`; do not assume names such as `Main.Run`.
- If the desired entrypoint is missing, add or fix the source module, run `xlflow push --json`, then rediscover macros before running.

Before designing a CLI-run macro, decide how inputs are supplied:

- Prefer `xlflow run <MacroName> --arg <type:value>` for user-provided paths, flags, and scalar settings.
- Use deterministic paths, environment variables, or configuration cells only when they are part of the project contract.
- Avoid UI prompts and active-window assumptions because unattended Excel automation cannot reliably answer them.
- When GUI behavior is required, keep the GUI entrypoint thin and extract the core logic into parameterized procedures that can run with `xlflow run --headless --arg`.

## Decision Flow

When the user asks to create or change VBA behavior:

1. Read `xlflow.toml` and relevant source files.
2. If the current source of truth is unclear, run `xlflow pull --json` before editing.
3. Edit `.bas`, `.cls`, or `.frm` source files.
4. Run `xlflow push --keepalive --json`.
5. Run `xlflow lint --json`.
6. Run `xlflow test --json` when tests exist.
7. If tests do not cover the behavior, run `xlflow macros --json`, then `xlflow run <qualified_name> --headless --keepalive --json` or `xlflow run <qualified_name> --trace --keepalive --json`.
8. Use `xlflow diff <before> <after> --json` when workbook state changes must be reviewed.

When the user reports a runtime failure:

1. Reproduce with `xlflow test --json` or `xlflow run <qualified_name> --trace --json`.
2. Inspect `error.code`, `error.phase`, VBA error metadata, and trace events before changing source.
3. Run `xlflow doctor --json` for setup phases such as `open_workbook`, `prepare_vbide`, or `inject_harness`.
4. Add targeted `XlflowLog` calls only around the suspected path, rerun, and keep the final trace noise low.
5. Patch the smallest relevant source area, then rerun the reproduction and broader verification.

## Command Usage

- Use `xlflow doctor` before blaming VBA when Excel automation fails before user code starts.
- Use `xlflow pull` to refresh editable source from the configured workbook.
- Use `xlflow push --keepalive` after source edits; it creates a backup before replacing VBA components.
- Use `xlflow lint` as the fast safety gate for generated VBA.
- Use `xlflow test` as the primary correctness signal when tests exist.
- Use `xlflow macros` to discover runnable macro entrypoints before guessing a `run` target.
- Use `xlflow inspect-gui --json` when a macro may require file pickers, message boxes, UserForms, or external process launches.
- Use `xlflow run --headless --keepalive` for repeatable automation; if it reports `gui_boundary_detected`, explain the boundary and either refactor the macro or rerun with `--interactive` when a human is available.
- Use `xlflow attach --active --json` before human-assisted sessions to confirm that the open Excel workbook matches `xlflow.toml`.
- Use `xlflow run --trace` when tests are absent, the macro mutates workbook state, or a runtime failure needs trace events.
- Use `xlflow diff` to summarize workbook and optional exported VBA differences.

## VBA Coding Rules

- Always use `Option Explicit`.
- Avoid `Select`, `Activate`, `ActiveSheet`, and unqualified `Range`.
- Prefer explicit workbook and worksheet references.
- Use `Long` instead of `Integer`.
- Keep procedures small and move business logic into testable standard modules.
- Restore `Application` state in cleanup paths.
- Use `On Error GoTo ErrHandler`; avoid broad `On Error Resume Next`.
- Do not pass object or array values to `AssertEquals`; compare scalar properties such as `Range.Value2`.
- Avoid UI prompts such as `Application.GetOpenFilename`, `Application.GetSaveAsFilename`, `Application.FileDialog`, `InputBox`, `UserForm.Show`, and modal `MsgBox` in macros that must run headlessly. Prefer `run --arg`, environment variables, configuration cells, or deterministic paths.
- Structure GUI-dependent macros with a human-facing entrypoint and a headless core:

```vb
Public Sub ImportData()
    Dim path As String
    path = PickImportFilePath()
    If path = "" Then Exit Sub
    ImportDataFromPath path
End Sub

Public Sub ImportDataFromPath(ByVal path As String)
    ' Core logic here
End Sub
```

## Keepalive Rules

Use `--keepalive --json` for long `xlflow push` and `xlflow run` commands. Keepalive heartbeat lines and the final `XLFLOW_DONE` marker are written to stderr so stdout remains valid JSON.

After starting a keepalive command, wait until the process exits and stderr contains a line beginning with `XLFLOW_DONE`. Do not begin the next workbook-dependent step just because stdout has not changed for a while.

Expected markers include `XLFLOW_DONE status=success command=push` and `XLFLOW_DONE status=failed command=run code=macro_timeout`.

## Trace Rules

Before using trace logging, run `xlflow trace inject --json` once for the configured workbook. In configured projects this also writes `src/modules/XlflowTrace.bas`, so later `xlflow push` keeps the trace module. If `xlflow run --trace --json` reports `trace_not_injected`, inject the trace module first, then rerun the macro.

When debugging, add `XlflowLog` calls at procedure entry, important branches, row or column counts, external paths, before and after destructive operations, and error handlers.

Keep high-level progress trace logs if they help future diagnosis. Remove noisy temporary logs before finalizing.

```vb
Call XlflowLog("start GenerateReport")
Call XlflowLog("lastRow=" & lastRow)
Call XlflowLog("finished GenerateReport")
```

## Failure Handling

If `xlflow test` fails, read the failing test name, module, VBA error number, description, and line. Patch the smallest relevant area, rerun the focused test first, then run the full test suite.

If `xlflow run` fails, inspect `error.code` and `error.phase`. `macro_not_found` means the entrypoint is missing or invalid; run `xlflow macros --json` and correct the target before changing user code. Setup phases such as `open_workbook`, `prepare_vbide`, and `inject_harness` usually indicate environment, configuration, or VBIDE access problems. `invoke_macro` points at the target macro or code it calls.

If `xlflow run --headless --json` fails with `gui_boundary_detected`, read `gui_boundaries` and do not retry the same command blindly. Either refactor the GUI boundary behind a parameterized core procedure, or switch to `xlflow run --interactive --json` only when a human is ready to operate Excel. If `macro_timeout` is returned, suspect an unresolved dialog, file picker, UserForm, or long-running loop.

If `xlflow run --trace` fails, read trace events from top to bottom, identify the last successful event, add targeted trace logs around the suspected block, and rerun. If the traced run fails with zero events, execution may have failed before reaching user `XlflowLog` calls; add an entry trace at the macro start or verify the macro target with `xlflow macros --json`.

If `xlflow lint` fails, fix lint findings directly in source files before rerunning `push`, `run`, or `test`.

## Final Response

Report:

- changed modules or files
- commands executed
- lint, test, macro, and diff results
- remaining risks or unverified conditions
