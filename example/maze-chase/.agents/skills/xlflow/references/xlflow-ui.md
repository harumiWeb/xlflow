# XlflowUI Dialog Reference

Load this reference when the task depends on `MsgBox`, `InputBox`, file pickers, or future interactive VBA helpers that should remain usable in headless `run`, `test`, or agent workflows.

## Core Rule

- Do not add raw `MsgBox`, `InputBox`, or file dialog calls in agent-authored VBA modules.
- Use `XlflowUI.MsgBox`, `XlflowUI.InputBox`, `XlflowUI.GetOpenFilename`, `XlflowUI.FileDialogOpen`, `XlflowUI.GetSaveAsFilename`, and `XlflowUI.FolderPicker` with stable dialog ids.
- Raw `VBA.Interaction.MsgBox`, `VBA.Interaction.InputBox`, `Application.GetOpenFilename`, `Application.GetSaveAsFilename`, and `Application.FileDialog` should appear only inside `XlflowUI.bas` or a clearly documented human-only adapter.
- When xlflow eventually supports more interactive helpers, extend `XlflowUI` rather than introducing raw VBA prompts directly in business logic.

## Runtime Contract

`XlflowUI` preserves one VBA call surface across interactive and unattended execution:

- `interactive`: delegates to native VBA or Excel dialog APIs
- `headless`, `ci`, `agent`, `test`: resolves scripted responses from xlflow-injected workbook markers

That means the same workbook code can be used by:

- an AI agent running `xlflow run --headless`
- an automated test running `xlflow test`
- a human opening the workbook in Excel normally

## Stable Dialog IDs

- Every `XlflowUI` dialog needs a stable id such as `confirm-save`, `customer-name`, or `overwrite-existing`.
- Ids must contain at least one ASCII letter or digit.
- xlflow normalizes ids to lowercase ASCII letters and digits joined by `_` for workbook-marker lookup.
- Do not create ids that normalize to the same value, such as `confirm save` and `confirm-save`.
- Keep ids semantic and stable across refactors so test fixtures and agent prompts remain reusable.
- `DefaultResponse` is the workbook-side fallback when no `--msgbox` value is supplied for a headless run or test.
- `DefaultValue` is the workbook-side fallback when no `--inputbox` value is supplied for a headless run or test.
- File dialog wrappers also accept `DefaultValue`; use `False` for cancel, one string for single-select, or a Variant string array for multi-select defaults.

## CLI Contract

Use repeated response flags on both `run` and `test`.

- `--msgbox <dialog-id=result>`
- `--inputbox <dialog-id=value>`
- `--filedialog <kind>:<dialog-id>=<value>`
- `--ui-stream`

Supported `--msgbox` results:

- `abort`
- `cancel`
- `ignore`
- `no`
- `ok`
- `retry`
- `yes`

Example:

```bash
xlflow run Main.Run --headless --msgbox confirm-save=yes --inputbox customer-name=alice --json
xlflow test --msgbox confirm-save=ok --inputbox customer-name=test-user --json
xlflow run Main.Run --headless --filedialog get-open:source-files=C:\temp\a.txt --filedialog get-open:source-files=C:\temp\b.txt --filedialog save-as:export-path=C:\temp\out.xlsx --json
```

Add `--ui-stream` when you need realtime visibility into how headless dialogs resolved:

```bash
xlflow run Main.Run --headless --msgbox confirm-save=yes --inputbox customer-name=alice --ui-stream --json
xlflow test --msgbox confirm-save=ok --inputbox customer-name=test-user --ui-stream --json
xlflow run Main.Run --headless --filedialog folder:export-dir=@cancel --ui-stream --json
```

Supported `--filedialog` kinds:

- `get-open`
- `file-open`
- `save-as`
- `folder`

Repeat the same `kind:id=value` flag to supply multi-select paths in order. Use `@cancel` to represent a cancelled file dialog.

`--ui-stream` writes realtime `XlflowUI` summaries to stderr, not stdout, so `--json` stdout remains machine-readable. Example stderr lines:

```text
xlflow: ui kind=msgbox id=confirm-save source=default result=yes
xlflow: ui kind=inputbox id=customer-name source=default value=[redacted]
xlflow: ui kind=file-open id=source-files source=scripted value=C:\temp\a.txt | C:\temp\b.txt
```

InputBox values are redacted by default in both the streamed stderr lines and the final JSON payload.

## Output Contract

When `--ui-stream` is enabled:

- stderr receives realtime `XlflowUI` event summaries
- stdout still contains the final human output or JSON envelope only
- final `run` / `test` results include top-level `ui.events`

When `--json` is not used, human-readable `run` and `test` output may also include a `UI` section summarizing the same events after execution.

## VBA Pattern

```vb
Dim decision As VbMsgBoxResult
Dim customerName As String
Dim sourceFiles As Variant

decision = XlflowUI.MsgBox("confirm-save", "Save workbook?", vbYesNo + vbQuestion, "Customer")
If decision <> vbYes Then Exit Sub

customerName = XlflowUI.InputBox("customer-name", "Customer name", "Customer", "")
sourceFiles = XlflowUI.GetOpenFilename("source-files", MultiSelect:=True)
```

Prefer ids that express the business decision, not UI wording.

## Design Rules For Agents

- Use `XlflowUI` only for true human interaction points.
- For machine-supplied values such as batch paths, feature flags, or automation-only parameters, prefer `xlflow run --arg`, config cells, or deterministic configuration instead of dialog wrappers.
- Use the file dialog wrappers when the same workbook flow genuinely needs both human file picking and scripted unattended execution.
- Keep one dialog id per distinct decision or input. Do not reuse one id for unrelated prompts.
- Keep dialog prompts thin. Business logic belongs after the wrapper returns.
- If the flow needs many fields, validation loops, or rich state, move that UX to a UserForm or worksheet-driven configuration instead of chaining many `InputBox` calls.

## Recommended Agent Workflow

1. Read `xlflow.toml` and relevant source modules.
2. If dialogs are involved, confirm the code uses `XlflowUI` rather than raw `MsgBox` / `InputBox` / file dialog calls.
3. Run `xlflow lint --json` and treat `VB007` on raw dialogs as a migration task.
4. Run `xlflow test --session --msgbox ... --inputbox ... --filedialog ... --ui-stream --json` when tests exist and realtime dialog visibility helps validation or debugging.
5. Otherwise run `xlflow run --headless --session --msgbox ... --inputbox ... --filedialog ... --ui-stream --json` when headless dialog behavior itself needs confirmation. Omit `--ui-stream` when only the final result matters.
6. Use `xlflow run --interactive` only when a human must operate non-`XlflowUI` GUI.

## Debugging Rule

If headless `XlflowUI` behavior is suspicious, rerun with the same `--msgbox` / `--inputbox` / `--filedialog` values plus `--ui-stream` before adding extra `XlflowDebug.Log` or other VBA logging. Compare:

- realtime stderr lines for the order of dialog resolution
- final `ui.events` for structured post-run inspection
- workbook-side `DefaultResponse` / `DefaultValue` expectations when `response_source=default`

## Future Extension Rule

If xlflow adds support for more interactive VBA functions later, follow the same pattern:

- add the wrapper under `XlflowUI`
- keep the interactive Excel behavior intact for humans
- add an unattended response transport through xlflow CLI/runtime injection
- teach lint/analyzer guidance to prefer the wrapper over the raw VBA function
