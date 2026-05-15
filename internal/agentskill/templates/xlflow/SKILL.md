---
name: xlflow
description: Use when Codex or another AI agent needs to edit, test, debug, or validate Excel VBA workbooks with xlflow. Provides the safe VBA development workflow for xlflow projects, including pull/push, lint, run, trace, test, diff, failure handling, and final reporting rules.
---

# xlflow Skill

## Purpose

Use xlflow as the proof loop for Excel VBA work. Do not treat generated VBA as complete until the workbook has been exported or inspected, source has been changed, changes have been imported, and the relevant macro or tests have been run.

Default safety rules for AI-agent work:

- Usually start with `xlflow session start` and stay in that session until the task is done.
- If it is unclear whether source files or the workbook are newer, start the session and run `xlflow pull --session --keepalive --json`.
- When `[vba].folders=true`, treat the filesystem layout under each configured `[src]` root as meaningful architecture. Nested directories map to Rubberduck-compatible `@Folder(...)` annotations during `push`.
- If `push` or `run` leaves the live session workbook unsaved, treat the live workbook as newer than disk until `xlflow save --json`.
- `xlflow inspect` reads the saved workbook file directly. Do not trust `inspect` to reflect unsaved live session changes until `xlflow save --json` has completed.
- Use `xlflow list forms --session --keepalive --json` when you need workbook UserForm names or expected `.frm` / `.frx` source paths without loading the form at runtime.
- Use `xlflow inspect form <FormName> --runtime|--designer|--both --session --keepalive --json` when you need structured UserForm state beyond saved worksheet snapshots. Runtime inspection executes against a temporary workbook copy of the current source state.
- Use `xlflow form snapshot <FormName> --out src/forms/specs/<FormName>.yaml --session --keepalive --json` when you need a persisted JSON/YAML spec for design review or future declarative UserForm workflows. This path is stricter than `inspect form --designer` because it executes an injected helper to recover concrete control types.
- Use `xlflow form build <spec> --session --keepalive --json` when you need to create a Designer-backed UserForm from a persisted spec under `src/forms/specs/`.
- Use `xlflow form build <spec> --session --overwrite --keepalive --json` when the intended workflow is to replace an existing UserForm from spec rather than mutate it in place.
- Use `xlflow form export-image <FormName> --out <path> --session --keepalive --json` when visual verification depends on the runtime-rendered UserForm rather than structured inspection alone.
- When a task depends on UserForm spec authoring or review, load [references/forms.md](references/forms.md) before choosing `inspect form`, `form snapshot`, or `form build`. It defines the persisted `xlflow.userform` schema, flat `controls` contract, overwrite safety rules, supported control types, and the best-effort versus observed-only fields that should not be treated as round-trip guarantees.
- Treat `src/forms/specs/*.yaml` as the canonical source-controlled artifact for UserForm design. Code-behind authority depends on `[userform].code_source`: new projects default to `sidecar`, where `src/forms/code/*.bas` is canonical, while imported projects default to `frm`, where embedded `.frm` code remains canonical until migration. `.frm` / `.frx` are generated Designer artifacts. After `xlflow form build`, xlflow now re-materializes those artifacts back into `src/forms/`, and `push` blocks before Excel opens when spec filename, `form.name`, `.frm` basename, or `.frm` `Attribute VB_Name` disagree.
- Use `xlflow export-image` when verification depends on rendered appearance rather than saved workbook cell/style snapshots alone.
- Use `xlflow edit --session` for temporary workbook-state setup, event triggering, and visual tuning when the change does not belong in production VBA yet.
- `xlflow run` returns structured compile diagnostics by default. Use `--gui-compile-errors` only when a human explicitly wants raw Excel/VBE compile dialogs.
- When the macro argument is omitted, `xlflow run` uses `project.entry` from `xlflow.toml`.

## Session Lifecycle

For normal AI-agent development tasks, use an explicit xlflow session from task start to task end:

1. Start with `xlflow session start` after reading `xlflow.toml` and resolving source-of-truth questions.
2. Matching sessions are auto-reused for `list forms`, `inspect form`, `form snapshot`, `form build`, `pull`, `push`, `macros`, `run`, `export-image`, `test`, `trace`, and `save` when the configured workbook path matches `.xlflow/session.json`; add `--session` when you want that reuse to be explicit.
3. Prefer `xlflow push --fast --session --no-save --keepalive --json` while iterating, and use `xlflow run --session --keepalive --json` or `xlflow run --headless --session --keepalive --json` when `project.entry` is the intended entrypoint because structured compile diagnostics are on by default.
4. Save with `xlflow save --json` before any disk-based verification step such as `xlflow inspect ...` when the live session workbook may be newer than disk.
5. End with `xlflow save --json` when workbook changes must persist, then always run `xlflow session stop`.

Use isolated non-session commands only for one-shot CI-style verification, release checks, suspicious session state, or when the user explicitly asks not to keep Excel open.

If `xlflow push --session --no-save` succeeds, or `xlflow run --session` completes without `--save` or `--save-as`, treat the live workbook as potentially newer than the `.xlsm` on disk until `xlflow save` runs.

## Standard Workflow

1. Inspect the project.
   - Read `xlflow.toml`.
   - Treat the configured source directories as authoritative when `xlflow.toml` exists and the user has not said the workbook contains newer VBA.
   - Run `xlflow doctor --keepalive --json` when Excel, COM, VBIDE, or macro execution looks suspicious.
   - Start `xlflow session start` for normal AI-agent development once the source of truth is clear.
   - Run `xlflow pull --session --keepalive --json` before editing when the workbook is the current source of truth.

2. Edit source files.
   - Prefer `.bas`, `.cls`, and `.frm` files under the configured source directories.
   - In folder mode, move files by directory when you want to change Rubberduck folder organization; `push` rewrites `@Folder(...)` from path in temporary import copies.
   - Do not edit binary workbooks directly unless the task is explicitly workbook-state only.

3. Import and check.
   - Prefer `xlflow push --fast --session --no-save --keepalive --json` after source edits while the session is running.
   - Use plain `xlflow push --keepalive --json` for CI-style verification, release checks, or when session state is suspect.
   - Run `xlflow lint --json` and fix reported issues before finalizing.

4. Execute behavior.
   - Prefer `xlflow test --session --keepalive --json`.
   - Use `xlflow test --filter <name> --session --keepalive --json` while iterating on one failing test.
   - If the macro entrypoint is unclear, run `xlflow macros --session --keepalive --json` before choosing a target.
   - If no tests exist and `project.entry` is the intended target, run `xlflow run --session --keepalive --json`.
   - Use `xlflow run <MacroName> --session --keepalive --json` when you need a non-default entrypoint.
   - Prefer `xlflow run --headless --session --keepalive --json` when `project.entry` is correct for unattended agent work that should still use the open session.
   - Use `xlflow run <MacroName> --headless --session --keepalive --json` when you need a non-default headless entrypoint.
   - Use `xlflow run <MacroName> --interactive --json` only when a human can operate Excel dialogs or forms.
   - Use `xlflow run <MacroName> --trace --session --json` when debugging runtime behavior or workbook mutation.

5. Inspect workbook results.
   - Use `xlflow list forms --session --keepalive --json` when the workbook contains UserForms and you need the authoritative form names before planning `inspect form`, snapshot, or source review work.
   - Use `xlflow inspect form <FormName> --runtime --session --keepalive --json` to read runtime state after form initialization or workbook-driven population logic; xlflow runs it against a temporary workbook copy so the source workbook state is preserved.
   - Use `xlflow inspect form <FormName> --designer --session --keepalive --json` to inspect design-time controls and geometry without running workbook VBA.
   - Use `xlflow form snapshot <FormName> --out src/forms/specs/<FormName>.yaml --session --keepalive --json` when you want that design-time state persisted as a reviewable JSON/YAML spec and you need concrete control types from the stricter helper path.
   - Use `xlflow form export-image <FormName> --out <path> --session --keepalive --json` when you need a PNG of the runtime-rendered UserForm for visual review.
   - Use `xlflow inspect form <FormName> --both --session --keepalive --json` when you need to compare designer state against runtime state in one pass.
   - Add `--initializer <MethodName>` to runtime or both mode when the form must be explicitly populated before inspection, such as workbook-scoped setup methods that mirror the visible UI.
   - Use `xlflow edit cell --sheet <name> --cell <A1> --value <text> --events on --session --keepalive --json` to prepare input cells and trigger `Worksheet_Change` handlers during a session.
   - Use `xlflow edit range --sheet <name> --range <A1:B2> --fill <#RRGGBB> --session --keepalive --json` or `--clear contents|formats|all` to reset workbook state between iterations.
   - Use `xlflow edit rows --sheet <name> --rows <1:31> --height <points> --session --keepalive --json` and `xlflow edit columns --sheet <name> --columns <A:AE> --width <chars> --session --keepalive --json` for visual tuning before `export-image`.
   - Use `xlflow inspect workbook --json` to confirm workbook path, active sheet metadata, and per-sheet used ranges.
   - Use `xlflow inspect sheets --json` to verify sheet creation/removal, visibility, row counts, and column counts.
   - Use `xlflow inspect range --sheet <name> --address <A1:F20> --json` when the expected output range is known.
   - Add `--include-style` when visual correctness depends on fill colors, borders, merged cells, row heights, or column widths.
   - Use `xlflow inspect used-range --sheet <name> --json` when the output bounds are unknown and you need the current data rectangle.
   - Use `xlflow inspect cell --sheet <name> --address <A1> --json` for targeted assertions on one cell.
   - Use `xlflow export-image --sheet <name> --range <A1:F20> --session --keepalive --json` when verification depends on the rendered sheet appearance, chart placement, fills, or layout.
   - Prefer global `--json` for agent parsing. Use `--format markdown` only when you intentionally want a compact human/LLM-facing table.
   - If the live session workbook is newer than disk, run `xlflow save --json` before relying on any `inspect` result.

6. Compare results.
   - Use `xlflow diff <before> <after> --json` for workbook state changes.
   - Add `--vba-before <dir> --vba-after <dir>` when exported source changes also need review.

7. Repeat until the command results prove the task.
   - Finish every normal AI-agent development task with `xlflow save --json` when workbook changes must persist, then `xlflow session stop`.

## Project Orientation

Before editing, decide what is authoritative:

- If `xlflow.toml` exists and source files are present, start a session, edit the configured source tree, and use `xlflow push --fast --session --json` during normal development.
- If the user says the workbook has the latest VBA, or source files are missing or stale, run `xlflow pull --session --keepalive --json` after starting the session, then edit source files.
- Do not mix direct workbook edits with source edits in the same task unless the requested change is workbook-state only and no VBA source change is needed.
- After `xlflow trace inject --keepalive --json`, remember that `XlflowTrace.bas` is generated xlflow support code. Do not rewrite it by hand unless the user is changing xlflow itself.
- Treat `xlflow inspect` as a disk snapshot reader, not a live Excel inspector. If the task depends on unsaved session changes, save first or use session-aware execution commands to reproduce the state again.

Before running a macro, decide the runnable entrypoint:

- Run `xlflow macros --session --keepalive --json` when the macro name is not already proven by tests, docs, prior command output, or `project.entry`.
- If `project.entry` is the intended entrypoint, prefer plain `xlflow run --session --keepalive --json` over repeating the macro name.
- Use a listed `qualified_name` from `xlflow macros --session --json`; do not assume names such as `Main.Run`.
- If the desired entrypoint is missing, add or fix the source module, run `xlflow push --fast --session --json` when a session is active, then rediscover macros before running.

Before designing a CLI-run macro, decide how inputs are supplied:

- Prefer `xlflow run <MacroName> --arg <type:value>` for user-provided paths, flags, and scalar settings.
- Use deterministic paths, environment variables, or configuration cells only when they are part of the project contract.
- Avoid UI prompts and active-window assumptions because unattended Excel automation cannot reliably answer them.
- When GUI behavior is required, keep the GUI entrypoint thin and extract the core logic into parameterized procedures that can run with `xlflow run --headless --arg`.

## Decision Flow

When the user asks to create or change VBA behavior:

1. Read `xlflow.toml` and relevant source files.
2. Start `xlflow session start` unless the task is a one-shot CI-style check, session state is suspect, or the user explicitly wants isolated commands.
3. If the current source of truth is unclear, run `xlflow pull --session --keepalive --json` before editing.
4. Edit `.bas`, `.cls`, or `.frm` source files.
5. Run `xlflow push --fast --session --no-save --keepalive --json`.
6. Run `xlflow lint --json`.
7. Run `xlflow test --session --keepalive --json` when tests exist.
8. If tests do not cover the behavior, run `xlflow macros --session --keepalive --json`, then `xlflow run --headless --session --keepalive --json` when `project.entry` is correct, or `xlflow run <qualified_name> --headless --session --keepalive --json` / `xlflow run <qualified_name> --trace --session --keepalive --json` for non-default entrypoints.
9. When workbook output matters, run `xlflow save --json` if needed, then inspect the result with `xlflow inspect workbook|sheets|range|used-range|cell --json`, or use `xlflow inspect form <FormName> --session --keepalive --json` for live UserForm state.
10. Run `xlflow save --json` when workbook changes must persist, then `xlflow session stop`.
11. Use `xlflow diff <before> <after> --json` when workbook state changes must be reviewed.

When the user reports a runtime failure:

1. Start `xlflow session start`, then reproduce with `xlflow test --session --keepalive --json` or `xlflow run <qualified_name> --trace --session --keepalive --json`.
2. Inspect `error.code`, `error.phase`, VBA error metadata, and trace events before changing source.
3. Run `xlflow doctor --keepalive --json` for setup phases such as `open_workbook`, `prepare_vbide`, or `inject_harness`.
4. Add targeted `XlflowLog` calls only around the suspected path, rerun, and keep the final trace noise low.
5. Patch the smallest relevant source area, then rerun the reproduction and broader verification.

## Command Usage

- Use `xlflow doctor --keepalive --json` before blaming VBA when Excel automation fails before user code starts.
- Use `xlflow session start` at the beginning of normal AI-agent development tasks, and use `xlflow session stop` before finalizing.
- Use `xlflow pull --session --keepalive --json` to refresh editable source from the configured workbook during a session.
- Use `xlflow push --fast --session --no-save --keepalive` after source edits during a session.
- Use plain `xlflow push --keepalive` when you need the safe isolated path with backup and save.
- Use `xlflow lint` as the fast safety gate for generated VBA.
- Use `xlflow test --session --keepalive --json` as the primary correctness signal when tests exist.
- Use `xlflow macros --session --keepalive --json` to discover runnable macro entrypoints before guessing a non-default `run` target.
- Use `xlflow inspect workbook --json` to confirm workbook-level metadata after save.
- Use `xlflow inspect sheets --json` to verify expected worksheet names, visibility, and lightweight used ranges.
- Use `xlflow inspect range --sheet <name> --address <A1:F20> --json` when the expected output rectangle is known.
- Use `xlflow inspect used-range --sheet <name> --json` when the output rectangle is unknown or may expand.
- Use `xlflow inspect cell --sheet <name> --address <A1> --json` for single-cell checks or precise assertions.
- Use `xlflow inspect form <FormName> --runtime|--designer|--both --session --keepalive --json` for UserForm inspection; prefer `--designer` for source/design audits, `--runtime` for populated runtime state from a temporary workbook copy, and `--both` when you need to compare them.
  - Use `xlflow form snapshot <FormName> --out src/forms/specs/<FormName>.yaml --session --keepalive --json` when the goal is a stable artifact for review, diff, or later declarative form work rather than one-off stdout inspection.
  - Use `xlflow form build src/forms/specs/<FormName>.yaml --session --keepalive --json` to materialize a new UserForm from a saved spec; add `--overwrite` when replacing an existing component intentionally.
- Use `xlflow form export-image <FormName> --out <path> --session --keepalive --json` when the question is visual fidelity of the runtime form, including layout after initialization.
- Use `xlflow inspect form <FormName> --runtime --initializer <MethodName> --session --keepalive --json` when the form requires an explicit initializer call before its visible state is meaningful.
- Use `xlflow save --json` before `inspect` whenever a session run or `push --session --no-save` may have left newer workbook state only in the live Excel instance.
- Treat `xlflow inspect workbook|sheets|range|used-range|cell` as disk-backed workbook readers, and `xlflow inspect form` as an Excel COM inspection command whose runtime mode executes against a temporary workbook copy derived from the current source workbook state.
- When UserForms are involved, treat `pull` / `push` / `save` warnings about partial `.frm` fidelity and unsaved session state as actionable; do not review UserForm diffs until the live workbook has been saved and re-pulled.
- Use `xlflow inspect-gui --json` when a macro may require file pickers, message boxes, UserForms, or external process launches.
- Use `xlflow edit --session` only for development-time workbook mutations; if the final application behavior depends on the styling or layout change, move that behavior back into reproducible VBA before finalizing.
- Use `xlflow run --headless --session --keepalive` for repeatable automation during normal development; if it reports `gui_boundary_detected`, explain the boundary and either refactor the macro or rerun with `--interactive` when a human is available.
- Plain `xlflow run --session --keepalive --json` already compiles first, uses `project.entry` when the macro argument is omitted, and returns structured compile diagnostics by default.
- Use `xlflow run --fast --session --keepalive --gui-compile-errors` only when a human explicitly accepts GUI compile dialogs and you intentionally want the direct fast path. Plain `xlflow run --direct` already opts out of default compile diagnostics automatically.
- Use `xlflow run --gui-compile-errors --interactive --json` only when a human explicitly wants raw compile dialogs instead of structured diagnostics.
- Matching workbook sessions auto-reuse on `list forms`, `inspect form`, `form snapshot`, `form build`, `pull`, `push`, `macros`, `run`, `export-image`, `test`, `trace`, and `save`; use explicit `--session` when you want that reuse to be deliberate and visible in the command line.
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

Use `--keepalive --json` for long Excel COM-backed commands, including `xlflow list forms`, `xlflow inspect form`, `xlflow form snapshot`, `xlflow form build`, `xlflow pull`, `xlflow push`, `xlflow macros`, `xlflow test`, `xlflow trace inject`, `xlflow run`, `xlflow export-image`, and workbook UI operations. Keepalive heartbeat lines and the final `XLFLOW_DONE` marker are written to stderr so stdout remains valid JSON.

After starting a keepalive command, wait until the process exits and stderr contains a line beginning with `XLFLOW_DONE`. Do not begin the next workbook-dependent step just because stdout has not changed for a while.

Expected markers include `XLFLOW_DONE status=success command=pull`, `XLFLOW_DONE status=success command=push`, and `XLFLOW_DONE status=failed command=run code=macro_timeout`.

## Trace Rules

Use `xlflow run --trace --session --keepalive --json` when you need trace events during normal development; xlflow can temporarily inject and revert the helper if it is missing. Use `xlflow trace enable --session --keepalive --json` when you want the helper persisted in the configured workbook and source tree. Use `xlflow trace status --session --json`, `xlflow trace disable --session --json`, and `xlflow trace clean --json` to inspect or remove trace state. `xlflow trace inject` is an older alias for `trace enable`.

Read the human output or top-level `trace.lifecycle` to tell whether the helper was temporary for one run or already persisted. If a traced run reports temporary helper injection but you want source-controlled tracing, follow with `xlflow trace enable --session --keepalive --json`.

When debugging, add `XlflowLog` calls at procedure entry, important branches, row or column counts, external paths, before and after destructive operations, and error handlers.

Keep high-level progress trace logs if they help future diagnosis. Remove noisy temporary logs before finalizing.

```vb
Call XlflowLog("start GenerateReport")
Call XlflowLog("lastRow=" & lastRow)
Call XlflowLog("finished GenerateReport")
```

## Windows PowerShell Checklist

When workbook code launches an external PowerShell process, separate xlflow's bridge host from the workbook-side host:

1. Check top-level `bridge.host` to see which PowerShell xlflow itself used.
2. Inspect the VBA command string or log the resolved executable from workbook code; it may be `powershell.exe` even when xlflow reports `pwsh.exe`, or the reverse.
3. If the issue looks like encoding or environment drift, standardize on one host before changing xlflow or VBA logic.

## Failure Handling

If `xlflow test` fails, read the failing test name, module, VBA error number, description, and line. Patch the smallest relevant area, rerun the focused test first, then run the full test suite.

If `xlflow run` fails, inspect `error.code`, `error.phase`, and any top-level `run_diagnostic`. `macro_not_found` means the entrypoint is missing or invalid; run `xlflow macros --session --keepalive --json` and correct the target before changing user code. Setup phases such as `open_workbook`, `prepare_vbide`, and `inject_harness` usually indicate environment, configuration, or VBIDE access problems. `invoke_macro` points at the target macro or code it calls. Plain `run` already includes compile-first diagnostics by default; do not switch to `--gui-compile-errors` unless a human explicitly wants GUI dialogs.

If `xlflow run --headless --session --json` fails with `gui_boundary_detected`, read `gui_boundaries` and do not retry the same command blindly. Either refactor the GUI boundary behind a parameterized core procedure, or switch to `xlflow run --interactive --json` only when a human is ready to operate Excel. If `macro_timeout` is returned, suspect an unresolved dialog, file picker, UserForm, or long-running loop.

If `xlflow run --diagnostic --session --json` fails with `vba_compile_failed`, inspect `run_diagnostic.kind`, `run_diagnostic.message`, `run_diagnostic.location`, and `run_diagnostic.nearby_code` before changing source. Treat dialog text as localized opaque text and fix the selected source location when available.

If `xlflow run --trace --session` fails, read trace events from top to bottom, identify the last successful event, add targeted trace logs around the suspected block, and rerun. If the traced run fails with zero events, execution may have failed before reaching user `XlflowLog` calls; add an entry trace at the macro start or verify the macro target with `xlflow macros --session --keepalive --json`.

If `xlflow lint` fails, fix lint findings directly in source files before rerunning `push`, `run`, or `test`.

Run `xlflow analyze --json` or `xlflow check --keepalive --json` before changing object-heavy VBA. Analyzer findings such as `VBA101`, `VBA102`, and `VBA103` usually mean a missing `Set` assignment.

If `xlflow inspect` does not show the workbook changes you expected, first decide whether disk is stale. A prior `xlflow push --session --no-save` or `xlflow run --session` can leave the live Excel workbook newer than the saved `.xlsm`; run `xlflow save --json` and inspect again before assuming the macro logic failed.

If a UserForm inspection looks wrong, first decide whether you need design-time state or runtime state. Use `xlflow inspect form <FormName> --designer` for control layout and static properties, and `--runtime` when labels, combo contents, or text boxes are populated by code. If runtime values are blank but the visible UI should be initialized, rerun with `--initializer <MethodName>` before changing source.

If the structure looks right but the visible layout is still uncertain, use `xlflow form export-image <FormName> --out <path> --session --keepalive --json` to capture the runtime-rendered form from a temporary workbook copy.

If `xlflow inspect used-range` is truncated, use the reported `returned_range` and warnings to choose a narrower follow-up `xlflow inspect range` query instead of blindly increasing prompt size.

If workbook correctness is visual and `inspect` still leaves ambiguity, use `xlflow export-image --sheet <name> --range <A1:F20> --session --keepalive --json` and inspect the generated PNG artifact under `.xlflow/artifacts/images/`.

## Final Response

Report:

- changed modules or files
- commands executed
- lint, test, macro, and diff results
- remaining risks or unverified conditions
