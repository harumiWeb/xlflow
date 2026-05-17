# xlflow trace

Manage VBA trace logging support and trace log cleanup.

## Usage

```bash
xlflow trace enable
xlflow trace status
xlflow trace disable
xlflow trace clean
xlflow trace inject
```

## Options and Arguments

| Option / argument | Description                                        | Default |
| ----------------- | -------------------------------------------------- | ------- |
| `enable`          | Enable trace helper support.                       | -       |
| `status`          | Report helper and log state.                       | -       |
| `disable`         | Remove trace helper support when safe.             | -       |
| `clean`           | Remove trace logs.                                 | -       |
| `inject`          | Inject helper into the active/session workbook.    | -       |
| `--session`       | Operate against the managed live session workbook. | false   |

## Examples

```bash
xlflow trace enable --json
xlflow run Main.Run --trace --json
xlflow trace clean --json
```

## Notes

::: tip
VBA code can call `XlflowLog` after trace support is available.
:::

::: warning
`trace disable` refuses unsafe removal when the helper appears modified or in use.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "trace status",
  "trace": { "enabled": true, "log": ".xlflow/trace.log" }
}
```

## Related

- [run](./run)
- [session](./session)
