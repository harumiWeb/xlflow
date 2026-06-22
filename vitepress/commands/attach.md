# xlflow attach

Validate that the active Excel workbook matches the configured workbook.

## Usage

```bash
xlflow attach --active
```

## Options and Arguments

| Option / argument | Description                                     | Default  |
| ----------------- | ----------------------------------------------- | -------- |
| `--active`        | Inspect the currently active workbook in Excel. | required |
| `--json`          | Emit a machine-readable match result.           | false    |

## Examples

```bash
xlflow attach --active
xlflow attach --active --json
```

## Notes

> [!IMPORTANT]
> `attach` is a safety check. It does not import source, change workbook state, or create a managed session.

`attach` uses the `.NET` bridge on Windows in `auto` mode. Deprecated `--bridge powershell` remains explicit opt-in for v0.15.0 only and emits a removal warning.

::: tip
Use this before manual Excel work when multiple workbooks are open.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "attach",
  "active_workbook": "Book.xlsm",
  "matches_config": true
}
```

## Related

- [session](./session)
- [doctor](./doctor)
