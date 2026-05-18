# xlflow inspect

Inspect saved workbook state or UserForm state.

## Usage

```bash
xlflow inspect workbook
xlflow inspect sheets
xlflow inspect range --sheet <sheet> --address <range>
xlflow inspect form <name> [--runtime|--designer|--both]
```

## Options and Arguments

| Option / argument                 | Description                                               | Default  |
| --------------------------------- | --------------------------------------------------------- | -------- |
| `workbook`                        | Return workbook-level metadata.                           | -        |
| `sheets`                          | List worksheets.                                          | -        |
| `range`                           | Read a worksheet range.                                   | -        |
| `used-range`                      | Read a worksheet used range.                              | -        |
| `cell`                            | Read one cell.                                            | -        |
| `form <name>`                     | Inspect UserForm state.                                   | -        |
| `--sheet <name>`                  | Worksheet for sheet-scoped inspection.                    | -        |
| `--address <range>`               | A1 address for range or cell inspection.                  | -        |
| `--max-rows <n>`                  | Limit rows returned by range or used-range inspection.    | 100      |
| `--max-cols <n>`                  | Limit columns returned by range or used-range inspection. | 30       |
| `--runtime / --designer / --both` | Choose UserForm inspection mode.                          | designer |

## Examples

```bash
xlflow inspect workbook --json
xlflow inspect range --sheet Result --address A1:F20 --json
xlflow inspect form CalendarForm --both --json
```

## Notes

::: warning
Most file-based inspection reads the saved workbook. Use session-aware commands when you need live dirty workbook state.
:::

> [!IMPORTANT]
> Runtime UserForm inspection may execute initializer code on a temporary workbook copy.

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "inspect range",
  "sheet": "Result",
  "address": "A1:F20",
  "values": [
    ["Name", "Amount"],
    ["A", 42]
  ]
}
```

## Related

- [export-image](./export-image)
- [list](./list)
- [JSON output](../reference/json-output)
