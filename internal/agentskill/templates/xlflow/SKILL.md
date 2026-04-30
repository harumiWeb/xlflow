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
   - Run `xlflow doctor --json` when Excel, COM, VBIDE, or macro execution looks suspicious.
   - Run `xlflow pull --json` before editing when the workbook is the current source of truth.

2. Edit source files.
   - Prefer `.bas`, `.cls`, and `.frm` files under the configured source directories.
   - Do not edit binary workbooks directly unless the task is explicitly workbook-state only.

3. Import and check.
   - Run `xlflow push --json` after source edits.
   - Run `xlflow lint --json` and fix reported issues before finalizing.

4. Execute behavior.
   - Prefer `xlflow test --json`.
   - Use `xlflow test --filter <name> --json` while iterating on one failing test.
   - If no tests exist, run the target macro with `xlflow run <MacroName> --json`.
   - Use `xlflow run <MacroName> --trace --json` when debugging runtime behavior or workbook mutation.

5. Compare results.
   - Use `xlflow diff <before> <after> --json` for workbook state changes.
   - Add `--vba-before <dir> --vba-after <dir>` when exported source changes also need review.

6. Repeat until the command results prove the task.

## Command Usage

- Use `xlflow doctor` before blaming VBA when Excel automation fails before user code starts.
- Use `xlflow pull` to refresh editable source from the configured workbook.
- Use `xlflow push` after source edits; it creates a backup before replacing VBA components.
- Use `xlflow lint` as the fast safety gate for generated VBA.
- Use `xlflow test` as the primary correctness signal when tests exist.
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

## Trace Rules

Before using trace logging, run `xlflow trace inject --json` once for the configured workbook. If `xlflow run --trace --json` reports `trace_not_injected`, inject the trace module first, then rerun the macro.

When debugging, add `XlflowLog` calls at procedure entry, important branches, row or column counts, external paths, before and after destructive operations, and error handlers.

Keep high-level progress trace logs if they help future diagnosis. Remove noisy temporary logs before finalizing.

```vb
Call XlflowLog("start GenerateReport")
Call XlflowLog("lastRow=" & lastRow)
Call XlflowLog("finished GenerateReport")
```

## Failure Handling

If `xlflow test` fails, read the failing test name, module, VBA error number, description, and line. Patch the smallest relevant area, rerun the focused test first, then run the full test suite.

If `xlflow run --trace` fails, read trace events from top to bottom, identify the last successful event, add targeted trace logs around the suspected block, and rerun.

If `xlflow lint` fails, fix lint findings directly in source files before rerunning `push`, `run`, or `test`.

## Final Response

Report:

- changed modules or files
- commands executed
- lint, test, macro, and diff results
- remaining risks or unverified conditions
