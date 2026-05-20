# XlflowUI

`XlflowUI` is xlflow's dialog-safe VBA wrapper layer for simple interactive prompts.

Use it when workbook code needs `MsgBox` or `InputBox` behavior that should work in both:

- normal Excel usage by a human
- headless `xlflow run`
- unattended `xlflow test`
- AI-agent-driven development workflows

## Why XlflowUI Exists

Raw `MsgBox` and `InputBox` calls block unattended Excel automation. `XlflowUI` keeps one VBA API surface while letting xlflow supply scripted responses in unattended modes.

- In `interactive` mode, `XlflowUI.MsgBox` and `XlflowUI.InputBox` delegate to the native VBA dialogs.
- In `headless`, `ci`, `agent`, and `test` modes, xlflow resolves scripted responses from runtime markers injected before VBA starts.

This lets one workbook support both autonomous agent workflows and normal Excel users.

`DefaultResponse` and `DefaultValue` are workbook-side fallbacks used when xlflow does not receive a scripted `--msgbox` or `--inputbox` value for that dialog id.

## VBA Contract

Recent `xlflow new` scaffolds include `src/modules/XlflowUI.bas` with these wrappers:

```vb
Public Function MsgBox(ByVal Id As String, ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "") As VbMsgBoxResult
Public Function InputBox(ByVal Id As String, ByVal Prompt As String, Optional ByVal Title As String = "", Optional ByVal Default As String = "") As String
```

Example:

```vb
Dim decision As VbMsgBoxResult
Dim customerName As String

decision = XlflowUI.MsgBox("confirm-save", "Save workbook?", vbYesNo + vbQuestion, "Orders")
If decision <> vbYes Then Exit Sub

customerName = XlflowUI.InputBox("customer-name", "Customer name", "Orders", "")
```

## Stable Dialog IDs

Every wrapper call needs a stable dialog id.

- Use semantic ids such as `confirm-save`, `customer-name`, or `overwrite-existing`.
- Ids must contain at least one ASCII letter or digit.
- xlflow normalizes ids to lowercase ASCII letters and digits joined by `_`.
- Avoid ids that normalize to the same value, such as `confirm save` and `confirm-save`.

Treat dialog ids as part of the workbook contract. Changing them breaks scripted responses and test fixtures.

## CLI Response Flags

`run` and `test` both accept repeated scripted dialog responses.

- `--msgbox <dialog-id=result>`
- `--inputbox <dialog-id=value>`

Supported `--msgbox` result values:

- `abort`
- `cancel`
- `ignore`
- `no`
- `ok`
- `retry`
- `yes`

Examples:

```bash
xlflow run Main.Run --headless --msgbox confirm-save=yes --inputbox customer-name=alice --json
xlflow test --msgbox confirm-save=ok --inputbox customer-name=test-user --json
```

## Lint And Analyzer Guidance

`xlflow lint` reports raw `MsgBox` and `InputBox` usage as `VB007` GUI-boundary warnings.

The intended remediation is not "never show dialogs". The intended remediation is:

1. replace raw `MsgBox` with `XlflowUI.MsgBox`
2. replace raw `InputBox` with `XlflowUI.InputBox`
3. provide `--msgbox` and `--inputbox` values for unattended runs

Only disable `VB007` with `[lint].forbid_interactive_input = false` for genuinely human-only projects. That setting does not make `run --headless` capable of answering raw dialogs.

## Design Guidance

- Use `XlflowUI` only for true user interaction points.
- For machine-supplied values such as file paths, feature flags, or batch parameters, prefer `xlflow run --arg`, configuration cells, or deterministic configuration instead of dialog wrappers.
- Keep dialogs thin and business-oriented. One dialog id should represent one decision or one scalar input.
- If the flow requires many fields, validation loops, or rich state, move that UX to a UserForm or worksheet-driven settings instead of chaining many `InputBox` calls.

## AI Agent Workflow

When an AI agent edits VBA that includes simple dialogs:

1. replace raw `MsgBox` / `InputBox` with `XlflowUI`
2. keep stable dialog ids in source control
3. validate with `xlflow lint --json`
4. run unattended verification with `xlflow test --msgbox ... --inputbox ... --json` or `xlflow run --headless --msgbox ... --inputbox ... --json`
5. keep `xlflow run --interactive` for truly human-operated UI only

## Future Interactive Helpers

xlflow is expected to support more interactive helper functions over time. Follow the same pattern for those additions:

- add the wrapper under `XlflowUI`
- preserve the normal interactive Excel behavior
- add an unattended response path through xlflow runtime injection
- teach lint/analyzer guidance to point users toward the wrapper instead of the raw VBA API
