# xlflow diff

Compare workbook files and optionally exported VBA source trees.

## Usage

```bash
xlflow diff <before.xlsm> <after.xlsm> [--vba-before <dir>] [--vba-after <dir>]
```

## Options and Arguments

| Option / argument    | Description                               | Default  |
| -------------------- | ----------------------------------------- | -------- |
| `before.xlsm`        | Baseline workbook.                        | required |
| `after.xlsm`         | Workbook to compare.                      | required |
| `--vba-before <dir>` | Baseline exported VBA source directory.   | -        |
| `--vba-after <dir>`  | Comparison exported VBA source directory. | -        |
| `--json`             | Return structured diff counts and paths.  | false    |

## Examples

```bash
xlflow diff before.xlsm after.xlsm --json
xlflow diff before.xlsm after.xlsm --vba-before before/src --vba-after after/src --json
```

## Notes

::: important
A successful comparison can still report differences. Inspect the JSON summary instead of treating exit code `0` as no changes.
:::

::: tip
Use `pull` before `diff` when you need workbook and source changes in one review.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "diff",
  "summary": { "workbook_diffs": 2, "vba_diffs": 1, "total_diffs": 3 }
}
```

## Related

- [pull](./pull)
- [JSON output](../reference/json-output)
