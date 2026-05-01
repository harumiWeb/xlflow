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
   - Run `xlflow doctor --keepalive --json` when Excel, COM, VBIDE, or macro execution looks suspicious.
   - Run `xlflow pull --keepalive --json` before editing when the workbook is the current source of truth.
   - For normal development work, start an explicit session with `xlflow session start` after the source of truth is clear.

2. Edit source files.
   - Prefer `.bas`, `.cls`, and `.frm` files under the configured source directories.
   - Do not edit binary workbooks directly unless the task is explicitly workbook-state only.

3. Import and check.
   - Prefer `xlflow push --fast --session --no-save --keepalive --json` after source edits while the session is running.
   - Use plain `xlflow push --keepalive --json` for CI-style verification, release checks, or when session state is suspect.
   - Run `xlflow lint --json` and fix reported issues before finalizing.

4. Execute behavior.
   - Prefer `xlflow test --keepalive --json`.
   - Use `xlflow test --filter <name> --keepalive --json` while iterating on one failing test.
   - If the macro entrypoint is unclear, run `xlflow macros --keepalive --json` before choosing a target.
   - If no tests exist, run the target macro with `xlflow run <MacroName> --fast --session --keepalive --json` when a session is active.
   - Prefer `xlflow run <MacroName> --headless --session --keepalive --json` for unattended agent work that should still use the open session.
   - Use `xlflow run <MacroName> --interactive --json` only when a human can operate Excel dialogs or forms.
   - Use `xlflow run <MacroName> --trace --json` when debugging runtime behavior or workbook mutation.

5. Compare results.
   - Use `xlflow diff <before> <after> --json` for workbook state changes.
   - Add `--vba-before <dir> --vba-after <dir>` when exported source changes also need review.

6. Repeat until the command results prove the task.
   - Finish session-based work with `xlflow save --session --json` when workbook changes must persist, then `xlflow session stop`.

## Project Orientation

Before editing, decide what is authoritative:

- If `xlflow.toml` exists and source files are present, edit the configured source tree and use `xlflow push --fast --session --json` during normal development.
- If the user says the workbook has the latest VBA, or source files are missing or stale, run `xlflow pull --keepalive --json` first and then edit source files.
- Do not mix direct workbook edits with source edits in the same task unless the requested change is workbook-state only and no VBA source change is needed.
- After `xlflow trace inject --keepalive --json`, remember that `XlflowTrace.bas` is generated xlflow support code. Do not rewrite it by hand unless the user is changing xlflow itself.

Before running a macro, decide the runnable entrypoint:

- Run `xlflow macros --keepalive --json` when the macro name is not already proven by tests, docs, or prior command output.
- Use a listed `qualified_name` from `xlflow macros --json`; do not assume names such as `Main.Run`.
- If the desired entrypoint is missing, add or fix the source module, run `xlflow push --fast --session --json` when a session is active, then rediscover macros before running.

Before designing a CLI-run macro, decide how inputs are supplied:

- Prefer `xlflow run <MacroName> --arg <type:value>` for user-provided paths, flags, and scalar settings.
- Use deterministic paths, environment variables, or configuration cells only when they are part of the project contract.
- Avoid UI prompts and active-window assumptions because unattended Excel automation cannot reliably answer them.
- When GUI behavior is required, keep the GUI entrypoint thin and extract the core logic into parameterized procedures that can run with `xlflow run --headless --arg`.

## Decision Flow

When the user asks to create or change VBA behavior:

1. Read `xlflow.toml` and relevant source files.
2. If the current source of truth is unclear, run `xlflow pull --keepalive --json` before editing.
3. Start `xlflow session start` unless the task is a one-shot CI-style check or the user explicitly wants isolated commands.
4. Edit `.bas`, `.cls`, or `.frm` source files.
5. Run `xlflow push --fast --session --no-save --keepalive --json`.
6. Run `xlflow lint --json`.
7. Run `xlflow test --keepalive --json` when tests exist.
8. If tests do not cover the behavior, run `xlflow macros --keepalive --json`, then `xlflow run <qualified_name> --headless --session --keepalive --json` or `xlflow run <qualified_name> --trace --session --keepalive --json`.
9. Run `xlflow save --session --json` when workbook changes must persist, then `xlflow session stop`.
10. Use `xlflow diff <before> <after> --json` when workbook state changes must be reviewed.

When the user reports a runtime failure:

1. Reproduce with `xlflow test --keepalive --json` or `xlflow run <qualified_name> --trace --session --keepalive --json` when a session is active.
2. Inspect `error.code`, `error.phase`, VBA error metadata, and trace events before changing source.
3. Run `xlflow doctor --keepalive --json` for setup phases such as `open_workbook`, `prepare_vbide`, or `inject_harness`.
4. Add targeted `XlflowLog` calls only around the suspected path, rerun, and keep the final trace noise low.
5. Patch the smallest relevant source area, then rerun the reproduction and broader verification.

## Command Usage

- Use `xlflow doctor --keepalive --json` before blaming VBA when Excel automation fails before user code starts.
- Use `xlflow pull --keepalive --json` to refresh editable source from the configured workbook.
- Use `xlflow session start` before normal source edit / push / run loops, and use `xlflow session stop` before finalizing.
- Use `xlflow push --fast --session --no-save --keepalive` after source edits during a session.
- Use plain `xlflow push --keepalive` when you need the safe isolated path with backup and save.
- Use `xlflow lint` as the fast safety gate for generated VBA.
- Use `xlflow test --keepalive --json` as the primary correctness signal when tests exist.
- Use `xlflow macros --keepalive --json` to discover runnable macro entrypoints before guessing a `run` target.
- Use `xlflow inspect-gui --json` when a macro may require file pickers, message boxes, UserForms, or external process launches.
- Use `xlflow run --headless --session --keepalive` for repeatable automation during normal development; if it reports `gui_boundary_detected`, explain the boundary and either refactor the macro or rerun with `--interactive` when a human is available.
- Use `xlflow run --fast --session --keepalive` for argument-free, trace-disabled smoke checks during tight iteration.
- Use explicit `--session` on `push` and `run`; normal commands do not auto-attach to an existing session.
- Use `xlflow attach --active --keepalive --json` before human-assisted sessions to confirm that the open Excel workbook matches `xlflow.toml`.
- Use `xlflow run --trace --session` when tests are absent, the macro mutates workbook state, or a runtime failure needs trace events during a session.
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

Use `--keepalive --json` for long Excel COM-backed commands, including `xlflow pull`, `xlflow push`, `xlflow macros`, `xlflow test`, `xlflow trace inject`, `xlflow run`, and workbook UI operations. Keepalive heartbeat lines and the final `XLFLOW_DONE` marker are written to stderr so stdout remains valid JSON.

After starting a keepalive command, wait until the process exits and stderr contains a line beginning with `XLFLOW_DONE`. Do not begin the next workbook-dependent step just because stdout has not changed for a while.

Expected markers include `XLFLOW_DONE status=success command=pull`, `XLFLOW_DONE status=success command=push`, and `XLFLOW_DONE status=failed command=run code=macro_timeout`.

## Trace Rules

Use `xlflow run --trace --keepalive --json` when you need trace events; xlflow can temporarily inject and revert the helper if it is missing. Use `xlflow trace enable --keepalive --json` when you want the helper persisted in the configured workbook and source tree. Use `xlflow trace status --json`, `xlflow trace disable --json`, and `xlflow trace clean --json` to inspect or remove trace state. `xlflow trace inject` is an older alias for `trace enable`.

When debugging, add `XlflowLog` calls at procedure entry, important branches, row or column counts, external paths, before and after destructive operations, and error handlers.

Keep high-level progress trace logs if they help future diagnosis. Remove noisy temporary logs before finalizing.

```vb
Call XlflowLog("start GenerateReport")
Call XlflowLog("lastRow=" & lastRow)
Call XlflowLog("finished GenerateReport")
```

## Failure Handling

If `xlflow test` fails, read the failing test name, module, VBA error number, description, and line. Patch the smallest relevant area, rerun the focused test first, then run the full test suite.

If `xlflow run` fails, inspect `error.code`, `error.phase`, and any top-level `run_diagnostic`. `macro_not_found` means the entrypoint is missing or invalid; run `xlflow macros --keepalive --json` and correct the target before changing user code. Setup phases such as `open_workbook`, `prepare_vbide`, and `inject_harness` usually indicate environment, configuration, or VBIDE access problems. `invoke_macro` points at the target macro or code it calls.

If `xlflow run --headless --json` fails with `gui_boundary_detected`, read `gui_boundaries` and do not retry the same command blindly. Either refactor the GUI boundary behind a parameterized core procedure, or switch to `xlflow run --interactive --json` only when a human is ready to operate Excel. If `macro_timeout` is returned, suspect an unresolved dialog, file picker, UserForm, or long-running loop.

If `xlflow run --trace` fails, read trace events from top to bottom, identify the last successful event, add targeted trace logs around the suspected block, and rerun. If the traced run fails with zero events, execution may have failed before reaching user `XlflowLog` calls; add an entry trace at the macro start or verify the macro target with `xlflow macros --keepalive --json`.

If `xlflow lint` fails, fix lint findings directly in source files before rerunning `push`, `run`, or `test`.

Run `xlflow analyze --json` or `xlflow check --keepalive --json` before changing object-heavy VBA. Analyzer findings such as `VBA101`, `VBA102`, and `VBA103` usually mean a missing `Set` assignment.

## Final Response

Report:

- changed modules or files
- commands executed
- lint, test, macro, and diff results
- remaining risks or unverified conditions
