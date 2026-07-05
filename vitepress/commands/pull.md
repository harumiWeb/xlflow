# xlflow pull

Export workbook VBA components and form artifacts into configured source directories.

## Usage

```bash
xlflow pull [--session] [--formulas]
```

## Options and Arguments

| Option / argument | Description                                                            | Default |
| ----------------- | ---------------------------------------------------------------------- | ------- |
| `--session`       | Pull from the managed live session workbook.                           | false   |
| `--formulas`      | Also refresh worksheet formula snapshots under `formulas/` on success. | false   |
| `--json`          | Report exported files and warnings.                                    | false   |

## Examples

```bash
xlflow pull
xlflow pull --session --json
xlflow pull --formulas --json
```

## Notes

::: tip
Pull before editing if the workbook may contain newer VBA than the source tree.
:::

> [!IMPORTANT]
> UserForm Designer state and code-behind may be written to separate sidecar paths depending on project configuration.

`--formulas` runs the normal VBA pull first. If it succeeds, xlflow also extracts worksheet formulas, defined names, formula references, and parse summaries into `formulas/`.

Formula extraction reads the saved workbook file directly. When using `pull --session --formulas`, save the session first if formulas were changed only in the live workbook.

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "pull",
  "workbook": "Book.xlsm",
  "written": ["src/modules/Main.bas", "src/forms/specs/UserForm1.yaml"],
  "output": {
    "formulas": {
      "dir": "formulas",
      "manifest": "formulas/manifest.json",
      "sheet_count": 2,
      "formula_region_count": 5,
      "parse_status_summary": {
        "ok": 4,
        "partial": 1,
        "failed": 0
      },
      "defined_name_count": 2
    }
  }
}
```

## Related

- [push](./push)
- [diff](./diff)
- [project structure](../reference/project-structure)
