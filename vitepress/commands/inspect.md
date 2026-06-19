# xlflow inspect

Inspect saved workbook state or UserForm state.

## Usage

```bash
xlflow inspect workbook
xlflow inspect sheets
xlflow inspect range --sheet <sheet> --address <range>
xlflow inspect range --sheet <sheet> --address <range> --session
xlflow inspect form <name> [--runtime|--designer|--both]
xlflow inspect calls [--path <dir-or-file>] [--from <module-or-procedure>] [--to <name>]
xlflow inspect symbols [--path <dir-or-file>] [--include-private] [--include-labels] [--module <name>]
```

## Options and Arguments

| Option / argument                 | Description                                                                             | Default                  |
| --------------------------------- | --------------------------------------------------------------------------------------- | ------------------------ |
| `workbook`                        | Return workbook-level metadata.                                                         | -                        |
| `sheets`                          | List worksheets.                                                                        | -                        |
| `range`                           | Read a worksheet range.                                                                 | -                        |
| `used-range`                      | Read a worksheet used range.                                                            | -                        |
| `cell`                            | Read one cell.                                                                          | -                        |
| `form <name>`                     | Inspect UserForm state.                                                                 | -                        |
| `calls`                           | Inspect exported VBA source call sites.                                                 | -                        |
| `symbols`                         | Inspect exported VBA source symbols.                                                    | -                        |
| `--sheet <name>`                  | Worksheet for sheet-scoped inspection.                                                  | -                        |
| `--address <range>`               | A1 address for range or cell inspection.                                                | -                        |
| `--max-rows <n>`                  | Limit rows returned by range or used-range inspection.                                  | 100                      |
| `--max-cols <n>`                  | Limit columns returned by range or used-range inspection.                               | 30                       |
| `--session`                       | Inspect the live workbook in the managed xlflow session.                                | false                    |
| `--runtime / --designer / --both` | Choose UserForm inspection mode.                                                        | runtime                  |
| `--path <dir-or-file>`            | Source path for `inspect calls` or `inspect symbols`.                                   | configured `[src]` roots |
| `--from <module-or-procedure>`    | Filter `inspect calls` to one caller module or procedure.                               | -                        |
| `--to <name>`                     | Filter `inspect calls` to one callee text or name.                                      | -                        |
| `--include-members`               | Compatibility flag for `inspect calls`; members are included by default.                | false                    |
| `--include-builtins`              | Compatibility flag for `inspect calls`; built-in-looking calls are included by default. | false                    |
| `--include-private`               | Include private and local symbols.                                                      | false                    |
| `--include-labels`                | Include labels and numeric line labels.                                                 | false                    |
| `--module <name>`                 | Filter `inspect symbols` to one module.                                                 | -                        |

## Examples

```bash
xlflow inspect workbook --json
xlflow inspect range --sheet Result --address A1:F20 --json
xlflow inspect range --sheet Result --address A1:F20 --session --include-style --json
xlflow inspect form CalendarForm --both --json
xlflow inspect calls --json
xlflow inspect calls --from Main.Run --to BuildReport
xlflow inspect symbols --json
xlflow inspect symbols --path src/modules --include-private
```

## Notes

::: warning
Most `inspect` commands read the saved workbook by default. Add `--session` when you need the live workbook state that is still open and unsaved in Excel.
:::

`inspect calls` and `inspect symbols` are source-only. They parse exported `.bas`, `.cls`, and `.frm` files with tree-sitter-vba and do not open Excel.

`inspect calls` reports syntax-level call sites with caller module/procedure context, callee text, argument count, named arguments, source range, parse recovery state, and conservative resolution status. It does not perform full VBA name binding, COM type-library resolution, host object model inference, overload resolution, or dynamic dispatch resolution.

Human output groups calls by source file and caller:

```text
src/modules/Main.bas
Main.Run
  -> App.RunCore                  src/modules/Main.bas:5
```

> [!IMPORTANT]
> Runtime UserForm inspection may execute initializer code on a temporary workbook copy.

## JSON Output Examples

Successful `--json` output uses the xlflow envelope plus command-specific fields.

### Calls

```json
{
  "status": "ok",
  "command": "inspect",
  "inspect": {
    "target": "calls",
    "source": "tree_sitter_vba",
    "root": "src",
    "calls": [
      {
        "file": "src/modules/Main.bas",
        "module": "Main",
        "caller": {
          "name": "Run",
          "kind": "sub",
          "qualifiedName": "Main.Run"
        },
        "callee": {
          "text": "App.RunCore",
          "baseName": "RunCore",
          "receiver": "App",
          "member": "RunCore"
        },
        "arguments": {
          "count": 1,
          "named": []
        },
        "range": {
          "startLine": 5,
          "startColumn": 5,
          "endLine": 5,
          "endColumn": 29
        },
        "parse": {
          "hasError": false,
          "hasMissing": false
        },
        "resolution": {
          "status": "matched",
          "candidates": [
            {
              "qualifiedName": "App.RunCore",
              "kind": "sub",
              "file": "src/modules/App.bas",
              "line": 4
            }
          ]
        }
      }
    ],
    "summary": {
      "files": 16,
      "calls": 799,
      "matched": 374,
      "unresolved": 132,
      "ambiguous": 27,
      "external": 12,
      "builtinLike": 205,
      "memberCalls": 49,
      "parseErrors": 0,
      "missingNodes": 0
    }
  }
}
```

### Symbols

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
