# xlflow edit

Mutate a live session workbook for setup, smoke tests, and visual tuning.

## Usage

```bash
xlflow edit cell --sheet <sheet> --cell <addr> --value <value> --session
xlflow edit range --sheet <sheet> --range <addr> --fill <color> --session
xlflow edit rows --sheet <sheet> --rows <rows> --height <points> --session
xlflow edit columns --sheet <sheet> --columns <cols> --width <points> --session
```

## Options and Arguments

| Option / argument      | Description                                        | Default                |
| ---------------------- | -------------------------------------------------- | ---------------------- | ----------------------------------------------- | ---- |
| `cell`                 | Edit a single cell value or formula.               | -                      |
| `range`                | Clear or fill a rectangular cell range.            | -                      |
| `rows`                 | Set row height on a worksheet.                     | -                      |
| `columns`              | Set column width on a worksheet.                   | -                      |
| `--sheet <name>`       | Target worksheet.                                  | required               |
| `--cell <addr>`        | Single A1 cell address.                            | required for `cell`    |
| `--range <addr>`       | A1 range address.                                  | required for `range`   |
| `--rows <selector>`    | Row selector such as `1` or `1:31`.                | required for `rows`    |
| `--columns <selector>` | Column selector such as `A` or `A:AE`.             | required for `columns` |
| `--value <value>`      | Scalar value to write.                             | -                      |
| `--formula <formula>`  | Formula to write.                                  | -                      |
| `--fill <color>`       | Fill color using `#RGB` or `#RRGGBB`.              | -                      |
| `--clear <mode>`       | Clear `contents`, `formats`, or `all` for a range. | -                      |
| `--height <points>`    | Row height in points.                              | -                      |
| `--width <characters>` | Column width in Excel character units.             | -                      |
| `--events <keep        | on                                                 | off>`                  | Control events for cell value or formula edits. | keep |
| `--session`            | Edit the managed live workbook.                    | required               |

## Examples

```bash
xlflow edit cell --sheet Input --cell B2 --value ABC123 --session --json
xlflow edit cell --sheet Input --cell C2 --formula "=LEN(B2)" --session --json
xlflow edit range --sheet QR --range A1:AE31 --fill "#FFFFFF" --session --json
```

## Notes

> [!IMPORTANT]
> `edit` is session-oriented. Save explicitly when the mutation should persist to disk.

`edit cell|range|rows|columns` uses the `.NET` bridge on Windows in `auto` mode. Deprecated `--bridge powershell` remains explicit opt-in for v0.15.0 only and emits a removal warning.

::: warning
Treat edit payloads as workbook mutations. Use disposable sessions or backups for destructive experiments.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "edit cell",
  "mutation": { "sheet": "Input", "cell": "B2", "value": "ABC123" },
  "session": { "dirty": true }
}
```

## Related

- [session](./session)
- [inspect](./inspect)
- [save](./save)
