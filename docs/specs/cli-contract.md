# xlflow CLI Contract

## Scope

This spec defines the MVP command, configuration, JSON output, and exit-code contracts for xlflow.

xlflow is a Windows-first Go CLI that treats Excel VBA projects as source-controlled code. Excel operations use PowerShell and Excel COM. Non-Excel commands such as `init` and `lint` should remain testable without Excel installed.

## Commands

```text
xlflow new [workbook]
xlflow init <workbook>
xlflow doctor [--json]
xlflow pull [--json]
xlflow push [--json]
xlflow run [macro] [--json]
xlflow lint [--json]
```

`new` creates a fresh macro-enabled workbook under `build/` and scaffolds the same project layout as `init`. Without an argument it creates `build/Book.xlsm`; when the argument has no extension, `.xlsm` is appended. Any other extension is rejected because workbook creation always uses Excel macro-enabled format `52`.

`run` uses the positional macro argument when provided. Otherwise it uses `project.entry` from `xlflow.toml`.

## Configuration

The MVP only auto-discovers `xlflow.toml` from the current working directory. `vba.toml` is intentionally not supported.

```toml
[project]
name = "sample"
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"
visible = false
display_alerts = false

[src]
modules = "src/modules"
classes = "src/classes"
forms = "src/forms"
workbook = "src/workbook"

[lint]
require_option_explicit = true
forbid_select = true
forbid_activate = true
forbid_on_error_resume_next = true
```

## JSON Envelope

All JSON output uses a stable top-level envelope.

```json
{
  "status": "ok",
  "command": "lint",
  "error": null,
  "logs": []
}
```

`status` is either `ok` or `failed`. `error` is `null` on success and a structured object on failure.

Command-specific fields are added at the top level:

- `diagnostics` for `doctor`
- `workbook` and `backup` for Excel file commands
- `macro` for `run`
- `issues` for `lint`

## Exit Codes

- `0`: success
- `1`: user-code or validation failure, including lint findings and macro failure
- `2`: CLI argument or configuration error
- `3`: environment failure, including Excel, COM, VBIDE, PowerShell, and script execution failures

## Lint Rules

- `VB001`: missing `Option Explicit`
- `VB002`: `Select` usage
- `VB003`: `Activate` usage
- `VB004`: `On Error Resume Next` usage
- `VB005`: possible implicit `Variant`
- `VB006`: module-level `Public` variable usage
