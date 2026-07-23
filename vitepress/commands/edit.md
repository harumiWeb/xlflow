# xlflow edit

Mutate a live session workbook for setup, smoke tests, and visual tuning.

## Usage

```bash
xlflow edit cell --sheet <sheet> --cell <addr> --value <value> --session
xlflow edit range --sheet <sheet> --range <addr> --fill <color> --session
xlflow edit formula --sheet <sheet> --range <addr> --formula-r1c1 <formula> --session
xlflow edit rows --sheet <sheet> --rows <rows> --height <points> --session
xlflow edit columns --sheet <sheet> --columns <cols> --width <characters> --session
xlflow edit sheet add --name <sheet> --session
```

## Options and Arguments

| Option / argument          | Description                                                          | Default                            |
| -------------------------- | -------------------------------------------------------------------- | ---------------------------------- |
| `cell`                     | Edit a single cell value or formula.                                 | -                                  |
| `range`                    | Clear or fill a rectangular cell range.                              | -                                  |
| `formula`                  | Edit formulas across a rectangular range.                            | -                                  |
| `rows`                     | Set row height on a worksheet.                                       | -                                  |
| `columns`                  | Set column width on a worksheet.                                     | -                                  |
| `sheet add`                | Create a worksheet.                                                  | -                                  |
| `--sheet <name>`           | Target worksheet.                                                    | required                           |
| `--name <name>`            | Worksheet name to create.                                            | required for `sheet add`           |
| `--before <name>`          | Insert the new worksheet before an existing worksheet.               | -                                  |
| `--after <name>`           | Insert the new worksheet after an existing worksheet.                | -                                  |
| `--if-missing`             | Succeed without creating a duplicate when the worksheet exists.      | false                              |
| `--cell <addr>`            | Single A1 cell address.                                              | required for `cell`                |
| `--range <addr>`           | A1 range address.                                                    | required for `range` and `formula` |
| `--rows <selector>`        | Row selector such as `1` or `1:31`.                                  | required for `rows`                |
| `--columns <selector>`     | Column selector such as `A` or `A:AE`.                               | required for `columns`             |
| `--value <value>`          | Scalar value to write.                                               | -                                  |
| `--formula <formula>`      | A1-style formula to write.                                           | -                                  |
| `--formula-r1c1 <formula>` | R1C1-style formula to write across a range.                          | -                                  |
| `--fill <color>`           | Fill color using `#RGB` or `#RRGGBB`.                                | -                                  |
| `--clear <mode>`           | Clear `contents`, `formats`, or `all` for a range.                   | -                                  |
| `--height <points>`        | Row height in points.                                                | -                                  |
| `--width <characters>`     | Column width in Excel character units.                               | -                                  |
| `--events <keep\|on\|off>` | Control events for cell value/formula edits and range formula edits. | keep                               |
| `--calculate`              | Calculate the target range after formula editing.                    | false                              |
| `--session`                | Edit the managed live workbook.                                      | required                           |

## Examples

```bash
xlflow edit cell --sheet Input --cell B2 --value ABC123 --session --json
xlflow edit cell --sheet Input --cell C2 --formula "=LEN(B2)" --session --json
xlflow edit range --sheet QR --range A1:AE31 --fill "#FFFFFF" --session --json
xlflow edit formula --sheet Invoice --range D2:D1001 --formula-r1c1 "=RC[-2]*RC[-1]" --session --json
xlflow edit sheet add --name Config --if-missing --session --json
xlflow edit sheet add --name Dashboard --after Invoices --session --json
```

## Notes

> [!IMPORTANT]
> `edit` is session-oriented. Save explicitly when the mutation should persist to disk.

`edit formula --formula-r1c1` is the recommended way to apply repeated formula patterns from `xlflow formulas pull` region snapshots. `--formula` assigns A1-style formulas using Excel's normal range formula semantics.

`edit sheet add` appends the new worksheet after the last worksheet by default. Use exactly one of `--before` or `--after` for positioned insertion.

`edit cell|range|formula|rows|columns|sheet add` uses the `.NET` bridge on Windows in `auto` mode.

::: warning
Treat edit payloads as workbook mutations. Use disposable sessions or backups for destructive experiments.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "edit",
  "edit": {
    "kind": "formula",
    "sheet": "Invoice",
    "range": "D2:D1001",
    "formula_mode": "r1c1",
    "formula": "=RC[-2]*RC[-1]",
    "cells_updated": 1000,
    "calculated": false
  },
  "session": { "dirty": true }
}
```

Worksheet creation success:

```json
{
  "status": "ok",
  "command": "edit",
  "edit": {
    "kind": "sheet",
    "sheet": "Dashboard",
    "created": true,
    "after": "Invoices"
  },
  "session": { "dirty": true }
}
```

## Related

- [session](./session)
- [inspect](./inspect)
- [save](./save)

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow edit` when the task matches the command description above. For a goal-oriented workflow, start with the [How-to guides](../guides/) and return here for exact options.

## Prerequisites

Check the project configuration and run `xlflow doctor --json` before workbook-backed operations. Source-only commands can run without Excel; commands that read or mutate a workbook require Windows Excel and VBIDE access.

## What this command reads and changes

The command reads the inputs and configuration described in its syntax and examples. Treat source files, the saved workbook, and a live session as separate states; add `--session` when the live workbook is authoritative. Any mutation is reversible only when a backup or explicit session save boundary exists.

## Effect on source-of-truth state

Use `xlflow status --json` before and after the command. A source edit normally requires `push`; a workbook edit normally requires `pull`; a dirty live session requires `save --session` or an intentional discard.

## Common workflows

Combine this command with the relevant [source/workbook/session workflow](../concepts/workbook-session-source), and use `--json` in scripts and agent loops.

## Common failures

Read the structured `error.code`, exit code, and recovery metadata instead of scraping terminal text. The [symptom-oriented troubleshooting guide](../help/troubleshooting) maps installation, execution, session, VS Code, and WSL failures to recovery steps.
