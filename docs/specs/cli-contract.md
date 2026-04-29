# xlflow CLI Contract

## Scope

This spec defines the MVP command, configuration, JSON output, and exit-code contracts for xlflow.

xlflow is a Windows-first Go CLI that treats Excel VBA projects as source-controlled code. Excel operations use PowerShell and Excel COM. Non-Excel commands such as `init` and `lint` should remain testable without Excel installed.

## Commands

```text
xlflow [--json] new [workbook]
xlflow [--json] init <workbook>
xlflow [--json] doctor
xlflow [--json] pull
xlflow [--json] push
xlflow [--json] run [macro]
xlflow [--json] test [--filter <name>]
xlflow [--json] lint
```

`--json` is a persistent global flag and can be used with every command, including `new` and `init`.

`new` creates a fresh macro-enabled workbook under `build/` and scaffolds the same project layout as `init`. Without an argument it creates `build/Book.xlsm`; when the argument has no extension, `.xlsm` is appended. Any other extension is rejected because workbook creation always uses Excel macro-enabled format `52`.

`init` accepts an existing workbook path, copies that workbook into the new project's `build/<basename>` path, and records that project-local `build/...` path in `xlflow.toml` under `[excel].path` (for example `build/Sales.xlsx`).

`pull` exports standard modules, class modules, userforms, and workbook document modules into the configured source directories. Userforms may emit both `.frm` and `.frx` artifacts. Document modules are exported as source text suitable for linting and re-import.

`run` uses the positional macro argument when provided. Otherwise it uses `project.entry` from `xlflow.toml`.

`test` opens the configured workbook, discovers argument-free `Sub` procedures from the workbook VBIDE state, and runs procedures whose names start with `Test` or end with `_Test`. `--filter` uses exact procedure-name matching. Duplicate discovered test names, no discovered tests, missing filter targets, and VBA test failures are validation failures. Excel, COM, VBIDE, PowerShell, and script failures are environment failures.

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
detect_implicit_variant = true
forbid_public_module_fields = true
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
- `tests` for `test`
- `issues` for `lint`

`test` result objects contain `name`, `module`, `status`, `duration_ms`, and an optional `error`.

## Exit Codes

- `0`: success
- `1`: user-code or validation failure, including lint findings, macro failure, VBA test failure, no tests found, missing filter targets, and duplicate test names
- `2`: CLI argument or configuration error
- `3`: environment failure, including Excel, COM, VBIDE, PowerShell, and script execution failures

## VBA Test Rules

New and initialized projects include `src/modules/XlflowAssert.bas` with `AssertEquals expected, actual, [message]`. The helper is scalar-only: it compares normal scalar values, treats `Null` as equal only to `Null`, and raises a clear assertion error for object or array inputs. Compare object properties such as `Range.Value2` instead of passing object references.

Example:

```vb
Public Sub TestCreateReport()
    AssertEquals 10, Sheets("Result").Range("A1").Value
End Sub
```

## Lint Rules

- `VB001`: missing `Option Explicit`
- `VB002`: `Select` usage
- `VB003`: `Activate` usage
- `VB004`: `On Error Resume Next` usage
- `VB005`: possible implicit `Variant`
- `VB006`: module-level `Public` variable usage
