# GUI Boundary and Human-Assisted Run Spec

## Goal

Make xlflow safer for AI agents and CI when VBA contains Excel GUI operations. GUI calls are explicit boundaries, not hidden automation targets. Headless runs fail before Excel starts; interactive runs make human participation intentional.

## GUI Boundary Model

`internal/gui` scans configured `.bas`, `.cls`, and `.frm` source files under modules, classes, forms, workbook, and `tests`.

Each boundary contains:

- `file`
- `line`
- `kind`: `file_picker`, `modal_dialog`, `user_form`, `external_process`, or `message_pump`
- `symbol`
- `severity`: `interactive-only`
- `message`
- `suggestion`

Initial symbols:

- `Application.GetOpenFilename`
- `Application.GetSaveAsFilename`
- `Application.FileDialog`
- `InputBox`
- `MsgBox`
- `UserForm.Show`
- `.Show vbModal`
- `DoEvents`
- `Shell`
- `CreateObject("WScript.Shell").Popup`

## CLI Behavior

- `xlflow lint` keeps `VB007` and adds boundary metadata to JSON findings.
- `xlflow inspect-gui` reports top-level `gui_boundaries` without opening Excel.
- `xlflow doctor` includes GUI boundary summary diagnostics when source is available.
- `xlflow run --headless` fails with `gui_boundary_detected` before Excel starts if any boundary exists.
- `xlflow run --interactive` runs with Excel visible and alerts enabled.
- `xlflow run --timeout <duration>` defaults to `5m`; timeout failures return `macro_timeout`.
- `xlflow attach --active` validates that the active Excel workbook matches configured `excel.path`. It does not retarget `pull`, `push`, or `run`.

## VBA Design Guidance

GUI entrypoints should be thin wrappers. Core behavior should live in parameterized procedures that can be invoked with `xlflow run --headless --arg`.
