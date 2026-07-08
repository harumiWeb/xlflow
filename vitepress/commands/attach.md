# xlflow attach

Validate that the active Excel workbook matches the configured workbook.

::: warning Deprecated
`xlflow attach --active` is deprecated. Use `xlflow session attach` when you want xlflow commands to operate on an already-open workbook.
:::

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
> `attach` is a legacy safety check. It does not import source, change workbook state, or create a session for later commands.

`attach` uses the `.NET` bridge on Windows in `auto` mode.

::: tip
Use `xlflow session attach` before manual Excel work when you want `push`, `pull`, `run`, `inspect`, or `save` to target the workbook that is already open.
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
