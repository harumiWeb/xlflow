# xlflow inspect

Inspect saved workbook state or UserForm state.

## Usage

```bash
xlflow inspect workbook
xlflow inspect sheets
xlflow inspect range --sheet <sheet> --address <range>
xlflow inspect range --sheet <sheet> --address <range> --session
xlflow inspect form <name> [--runtime|--designer|--both]
xlflow inspect symbols [--path <dir-or-file>] [--include-private] [--include-labels] [--module <name>]
```

## Options and Arguments

| Option / argument                 | Description                                               | Default                  |
| --------------------------------- | --------------------------------------------------------- | ------------------------ |
| `workbook`                        | Return workbook-level metadata.                           | -                        |
| `sheets`                          | List worksheets.                                          | -                        |
| `range`                           | Read a worksheet range.                                   | -                        |
| `used-range`                      | Read a worksheet used range.                              | -                        |
| `cell`                            | Read one cell.                                            | -                        |
| `form <name>`                     | Inspect UserForm state.                                   | -                        |
| `symbols`                         | Inspect exported VBA source symbols.                      | -                        |
| `--sheet <name>`                  | Worksheet for sheet-scoped inspection.                    | -                        |
| `--address <range>`               | A1 address for range or cell inspection.                  | -                        |
| `--max-rows <n>`                  | Limit rows returned by range or used-range inspection.    | 100                      |
| `--max-cols <n>`                  | Limit columns returned by range or used-range inspection. | 30                       |
| `--session`                       | Inspect the live workbook in the managed xlflow session.  | false                    |
| `--runtime / --designer / --both` | Choose UserForm inspection mode.                          | runtime                  |
| `--path <dir-or-file>`            | Source path for `inspect symbols`.                        | configured `[src]` roots |
| `--include-private`               | Include private and local symbols.                        | false                    |
| `--include-labels`                | Include labels and numeric line labels.                   | false                    |
| `--module <name>`                 | Filter `inspect symbols` to one module.                   | -                        |

## Examples

```bash
xlflow inspect workbook --json
xlflow inspect range --sheet Result --address A1:F20 --json
xlflow inspect range --sheet Result --address A1:F20 --session --include-style --json
xlflow inspect form CalendarForm --both --json
xlflow inspect symbols --json
xlflow inspect symbols --path src/modules --include-private
```

## Notes

::: warning
Most `inspect` commands read the saved workbook by default. Add `--session` when you need the live workbook state that is still open and unsaved in Excel.
:::

`inspect symbols` is source-only. It parses exported `.bas`, `.cls`, and `.frm` files with tree-sitter-vba and does not open Excel.

> [!IMPORTANT]
> Runtime UserForm inspection may execute initializer code on a temporary workbook copy.

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "inspect",
  "inspect": {
    "target": "symbols",
    "source": "tree_sitter_vba",
    "root": "src",
    "files": [
      {
        "path": "src/modules/Main.bas",
        "moduleName": "Main",
        "moduleKind": "standard",
        "parse": { "hasError": false, "hasMissing": false },
        "symbols": [
          {
            "name": "Run",
            "kind": "sub",
            "module": "Main",
            "file": "src/modules/Main.bas",
            "startLine": 4,
            "startColumn": 1
          }
        ]
      }
    ],
    "summary": { "files": 1, "symbols": 1, "parseErrors": 0, "missingNodes": 0 }
  }
}
```

## Related

- [export-image](./export-image)
- [list](./list)
- [JSON output](../reference/json-output)
