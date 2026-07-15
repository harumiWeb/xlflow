# XlflowUI

`XlflowUI` is xlflow's dialog-safe VBA wrapper layer for simple interactive prompts and file-selection flows.

Use it when workbook code needs dialog behavior that should work in both:

- normal Excel usage by a human
- headless `xlflow run`
- unattended `xlflow test`
- AI-agent-driven development workflows

## Why XlflowUI Exists

Raw `MsgBox`, `InputBox`, and file dialog calls block unattended Excel automation. `XlflowUI` keeps one VBA API surface while letting xlflow supply scripted responses in unattended modes.

- In `interactive` mode, `XlflowUI` delegates to the native VBA or Excel dialogs.
- In `headless`, `ci`, `agent`, and `test` modes, xlflow resolves scripted responses from runtime markers injected before VBA starts.

This lets one workbook support both autonomous agent workflows and normal Excel users.

`DefaultResponse` and `DefaultValue` are workbook-side fallbacks used when xlflow does not receive a scripted `--msgbox` or `--inputbox` value for that dialog id.

`Default` is the native interactive prompt default shown only when a human is using Excel interactively. `DefaultValue` is the separate headless fallback used only when xlflow cannot find a scripted `--inputbox` value.

## VBA Contract

Recent `xlflow new` scaffolds include `src/modules/Xlflow/XlflowUI.bas` with these wrappers:

```vb
Public Function MsgBox(ByVal Id As String, ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "", Optional ByVal DefaultResponse As String = "") As VbMsgBoxResult
Public Function InputBox(ByVal Id As String, ByVal Prompt As String, Optional ByVal Title As String = "", Optional ByVal Default As String = "", Optional ByVal DefaultValue As String = "") As String
Public Function GetOpenFilename(ByVal Id As String, Optional ByVal FileFilter As String = "", Optional ByVal FilterIndex As Long = 1, Optional ByVal Title As String = "", Optional ByVal ButtonText As String = "", Optional ByVal MultiSelect As Boolean = False, Optional ByVal DefaultValue As Variant) As Variant
Public Function FileDialogOpen(ByVal Id As String, Optional ByVal Title As String = "", Optional ByVal ButtonText As String = "", Optional ByVal MultiSelect As Boolean = False, Optional ByVal DefaultValue As Variant) As Variant
Public Function GetSaveAsFilename(ByVal Id As String, Optional ByVal InitialFileName As String = "", Optional ByVal FileFilter As String = "", Optional ByVal FilterIndex As Long = 1, Optional ByVal Title As String = "", Optional ByVal ButtonText As String = "", Optional ByVal DefaultValue As Variant) As Variant
Public Function FolderPicker(ByVal Id As String, Optional ByVal Title As String = "", Optional ByVal ButtonText As String = "", Optional ByVal InitialPath As String = "", Optional ByVal DefaultValue As Variant) As Variant
```

Example:

```vb
Dim decision As VbMsgBoxResult
Dim customerName As String

decision = XlflowUI.MsgBox("confirm-save", "Save workbook?", vbYesNo + vbQuestion, "Orders")
If decision <> vbYes Then Exit Sub

customerName = XlflowUI.InputBox("customer-name", "Customer name", "Orders", "")
```

File dialog example:

```vb
Dim sourceFiles As Variant
Dim exportPath As Variant

sourceFiles = XlflowUI.GetOpenFilename("source-files", MultiSelect:=True)
exportPath = XlflowUI.GetSaveAsFilename("export-path")
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
- `--filedialog <kind>:<dialog-id>=<value>`
- `--ui-stream`

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
xlflow run Main.Run --headless --filedialog get-open:source-files=C:\temp\a.txt --filedialog get-open:source-files=C:\temp\b.txt --filedialog save-as:export-path=C:\temp\out.xlsx --json
```

Add `--ui-stream` when you want realtime visibility into how headless dialogs resolved:

```bash
xlflow run Main.Run --headless --msgbox confirm-save=yes --inputbox customer-name=alice --ui-stream --json
xlflow test --msgbox confirm-save=ok --inputbox customer-name=test-user --ui-stream --json
xlflow run Main.Run --headless --filedialog folder:export-dir=@cancel --ui-stream --json
```

Supported `--filedialog` kinds:

- `get-open` for `XlflowUI.GetOpenFilename`
- `file-open` for `XlflowUI.FileDialogOpen`
- `save-as` for `XlflowUI.GetSaveAsFilename`
- `folder` for `XlflowUI.FolderPicker`

File dialog values use these rules:

- Repeat the same `kind:id=value` flag to supply multiple selected paths in order.
- Use `@cancel` to represent a cancelled dialog.
- `GetOpenFilename` and `FileDialogOpen` return a Variant string array when `MultiSelect=True`.
- Non-multi-select wrappers return a single string or `False` on cancel.

`--ui-stream` writes realtime `XlflowUI` summaries to stderr, not stdout, so `--json` stdout stays machine-readable. Example stderr lines:

```text
xlflow: ui kind=msgbox id=confirm-save source=default result=yes
xlflow: ui kind=inputbox id=customer-name source=default value=[redacted]
xlflow: ui kind=file-open id=source-files source=scripted value=C:\temp\a.txt | C:\temp\b.txt
```

InputBox values are redacted by default in both streamed stderr lines and final result payloads.

## Output Contract

When `--ui-stream` is enabled:

- stderr receives realtime `XlflowUI` event summaries
- stdout still contains only the final human output or JSON envelope
- final `run` / `test` results include top-level `ui.events`

Without `--json`, human-readable `run` and `test` output may also include a `UI` section summarizing those same events after execution.

## Lint And Analyzer Guidance

`xlflow lint` reports raw `MsgBox`, `InputBox`, and file dialog usage as `VB007` GUI-boundary warnings.

The intended remediation is not "never show dialogs". The intended remediation is:

1. replace raw `MsgBox` with `XlflowUI.MsgBox`
2. replace raw `InputBox` with `XlflowUI.InputBox`
3. replace raw file picker calls with the matching `XlflowUI` file dialog wrapper
4. provide `--msgbox`, `--inputbox`, and `--filedialog` values for unattended runs

Only disable `VB007` with `[lint].forbid_interactive_input = false` for genuinely human-only projects. That setting does not make `run --headless` capable of answering raw dialogs.

When `XlflowUI.bas` is present, `push` also rejects bare `MsgBox` and `InputBox` calls before Excel opens. This is separate from `VB007`: unqualified names can bind to `XlflowUI.MsgBox` / `XlflowUI.InputBox` instead of the VBA built-ins. Use `XlflowUI` wrappers by default. If a module intentionally needs the native human-only dialog, write `VBA.Interaction.MsgBox` or `VBA.Interaction.InputBox` explicitly.

## Design Guidance

- Use `XlflowUI` only for true user interaction points.
- For machine-supplied values such as batch paths, feature flags, or automation-only inputs, prefer `xlflow run --arg`, configuration cells, or deterministic configuration instead of dialog wrappers.
- Use the file dialog wrappers when the same workbook flow genuinely needs both human file picking and scripted unattended execution.
- Keep dialogs thin and business-oriented. One dialog id should represent one decision or one scalar input.
- If the flow requires many fields, validation loops, or rich state, move that UX to a UserForm or worksheet-driven settings instead of chaining many `InputBox` calls.

## AI Agent Workflow

When an AI agent edits VBA that includes simple dialogs:

1. replace raw `MsgBox` / `InputBox` / file dialog calls with `XlflowUI`
2. keep stable dialog ids in source control
3. validate with `xlflow lint --json`
4. run unattended verification with `xlflow test --msgbox ... --inputbox ... --filedialog ... --ui-stream --json` or `xlflow run --headless --msgbox ... --inputbox ... --filedialog ... --ui-stream --json` when realtime dialog visibility matters; omit `--ui-stream` when only the final result matters
5. keep `xlflow run --interactive` for truly human-operated UI only

## Debugging Headless Dialogs

If headless `XlflowUI` behavior is unclear, rerun with the same `--msgbox` / `--inputbox` / `--filedialog` values plus `--ui-stream` before adding extra `XlflowDebug.Log` or other VBA logging.

- Use realtime stderr lines to confirm dialog order and response source.
- Use final `ui.events` for structured post-run inspection.
- If `response_source=default`, verify the workbook-side `DefaultResponse` / `DefaultValue` behavior before changing CLI fixtures.

## Future Interactive Helpers

xlflow is expected to support more interactive helper functions over time. Follow the same pattern for those additions:

- add the wrapper under `XlflowUI`
- preserve the normal interactive Excel behavior
- add an unattended response path through xlflow runtime injection
- teach lint/analyzer guidance to point users toward the wrapper instead of the raw VBA API
