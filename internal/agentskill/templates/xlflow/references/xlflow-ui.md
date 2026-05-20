# XlflowUI Dialog Reference

Load this reference when the task depends on `MsgBox`, `InputBox`, or future interactive VBA helpers that should remain usable in headless `run`, `test`, or agent workflows.

## Core Rule

- Do not add raw `MsgBox` or `InputBox` in agent-authored VBA modules.
- Use `XlflowUI.MsgBox` and `XlflowUI.InputBox` with stable dialog ids.
- Raw `VBA.Interaction.MsgBox` and `VBA.Interaction.InputBox` should appear only inside `XlflowUI.bas` or a clearly documented human-only adapter.
- When xlflow eventually supports more interactive helpers, extend `XlflowUI` rather than introducing raw VBA prompts directly in business logic.

## Runtime Contract

`XlflowUI` preserves one VBA call surface across interactive and unattended execution:

- `interactive`: delegates to native `VBA.Interaction.MsgBox` / `VBA.Interaction.InputBox`
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

## CLI Contract

Use repeated response flags on both `run` and `test`.

- `--msgbox <dialog-id=result>`
- `--inputbox <dialog-id=value>`

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
```

## VBA Pattern

```vb
Dim decision As VbMsgBoxResult
Dim customerName As String

decision = XlflowUI.MsgBox("confirm-save", "Save workbook?", vbYesNo + vbQuestion, "Customer")
If decision <> vbYes Then Exit Sub

customerName = XlflowUI.InputBox("customer-name", "Customer name", "Customer", "")
```

Prefer ids that express the business decision, not UI wording.

## Design Rules For Agents

- Use `XlflowUI` only for true human interaction points.
- For machine-supplied values such as file paths, feature flags, or batch parameters, prefer `xlflow run --arg`, config cells, or deterministic configuration instead of dialog wrappers.
- Keep one dialog id per distinct decision or input. Do not reuse one id for unrelated prompts.
- Keep dialog prompts thin. Business logic belongs after the wrapper returns.
- If the flow needs many fields, validation loops, or rich state, move that UX to a UserForm or worksheet-driven configuration instead of chaining many `InputBox` calls.

## Recommended Agent Workflow

1. Read `xlflow.toml` and relevant source modules.
2. If dialogs are involved, confirm the code uses `XlflowUI` rather than raw `MsgBox` / `InputBox`.
3. Run `xlflow lint --json` and treat `VB007` on raw dialogs as a migration task.
4. Run `xlflow test --session --msgbox ... --inputbox ... --json` when tests exist.
5. Otherwise run `xlflow run --headless --session --msgbox ... --inputbox ... --json`.
6. Use `xlflow run --interactive` only when a human must operate non-`XlflowUI` GUI.

## Future Extension Rule

If xlflow adds support for more interactive VBA functions later, follow the same pattern:

- add the wrapper under `XlflowUI`
- keep the interactive Excel behavior intact for humans
- add an unattended response transport through xlflow CLI/runtime injection
- teach lint/analyzer guidance to prefer the wrapper over the raw VBA function
