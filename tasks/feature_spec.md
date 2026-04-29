# xlflow MVP Feature Spec

## Goal

Implement the MVP command set from `docs/design.md` as a Windows-first, agent-ready VBA development CLI.

## Commands

- `xlflow init <workbook>`: scaffold project directories, copy the workbook to `build/`, write `xlflow.toml`, and write `prompts/agent.md`.
- `xlflow doctor [--json]`: diagnose Excel COM availability, workbook open support, and VBIDE trust access.
- `xlflow pull [--json]`: export workbook VBA components into configured source directories.
- `xlflow push [--json]`: back up current workbook VBA components, then import source files into the workbook.
- `xlflow run [macro] [--json]`: run the provided macro or `project.entry`.
- `xlflow lint [--json]`: lint `.bas`, `.cls`, and `.frm` files under configured source directories and `tests`.

## Configuration Types

`xlflow.toml` is the only MVP configuration filename.

```go
type Config struct {
    Project ProjectConfig
    Excel   ExcelConfig
    Src     SourceConfig
    Lint    LintConfig
}
```

Defaults:

- `excel.path`: `build/Book.xlsm`
- `excel.visible`: `false`
- `excel.display_alerts`: `false`
- `src.modules`: `src/modules`
- `src.classes`: `src/classes`
- `src.forms`: `src/forms`
- `src.workbook`: `src/workbook`
- lint booleans: `true`

## Result Types

All JSON output follows the stable envelope in `docs/specs/cli-contract.md`.

Errors are classified into:

- validation failures: exit code `1`
- config and argument errors: exit code `2`
- environment/script/Excel failures: exit code `3`

## PowerShell Bridge

Go executes scripts with:

```text
powershell -NoProfile -ExecutionPolicy Bypass -File <script>
```

Scripts must write only JSON to stdout and must release COM objects in cleanup paths.

## Lint Issues

Each lint issue includes `code`, `severity`, `file`, `line`, and `message`.
