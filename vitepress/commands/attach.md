# xlflow attach

Validate that the active Excel workbook matches the configured workbook.

## Usage

```bash
xlflow attach --active [--keepalive] [--keepalive-interval <duration>]
```

## Options and Arguments

| Option / argument | Description                                     | Default  |
| ----------------- | ----------------------------------------------- | -------- |
| `--active`        | Inspect the currently active workbook in Excel. | required |
| `--keepalive`     | Keep the bridge process alive.                  | false    |
| `--json`          | Emit a machine-readable match result.           | false    |

## Examples

```bash
xlflow attach --active
xlflow attach --active --json
```

## Notes

> [!IMPORTANT]
> `attach` is a safety check. It does not import source, change workbook state, or create a managed session.

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
