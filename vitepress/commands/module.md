# xlflow module

Create, remove, rename, or install VBA module source files in an existing project.

## Usage

```bash
xlflow module new <name> --type standard
xlflow module new <name> --type class
xlflow module remove <module-name>
xlflow module rename <old-name> <new-name>
xlflow module install [--push]
```

## Options and Arguments

| Option / argument              | Description                                                                                | Default  |
| ------------------------------ | ------------------------------------------------------------------------------------------ | -------- |
| `new <name>`                   | Create a new standard or class module source file.                                         | -        |
| `--type`                       | Required for `new`; must be `standard` or `class`.                                         | required |
| `remove <module-name>`         | Remove a standard, class, or UserForm source module from the project.                      | -        |
| `rename <old-name> <new-name>` | Rename a standard, class, or UserForm source module in the project.                        | -        |
| `install`                      | Install the bundled helper modules into the configured module source root.                 | required |
| `--push`                       | Push the installed helper modules into the configured workbook after writing source files. | false    |

## Examples

```bash
xlflow module new InvoiceProcessor --type standard
xlflow module new InvoiceService --type class --json
xlflow module remove InvoiceAggregator
xlflow module rename OldInvoiceService InvoiceService --json
xlflow module install
xlflow module install --push --json
```

## Notes

::: tip
`module new --type standard` writes `[src].modules/<Name>.bas`; `module new --type class` writes `[src].classes/<Name>.cls`.
:::

::: tip
`module install` writes `XlflowAssert.bas`, `XlflowRuntime.bas`, `XlflowUI.bas`, and `XlflowDebug.bas` into the configured `[src].modules` root.
:::

::: warning
`module new` and `module install` refuse to overwrite existing source files. Remove or rename colliding files before retrying.
:::

::: warning
`module remove` and `module rename` change source files only. Run `xlflow push` afterward to apply the change to the workbook.
:::

::: tip
UserForm remove/rename handles related source artifacts together: `.frm`, `.frx`, `src/forms/code/<Name>.bas`, and `src/forms/specs/<Name>.yaml|yml|json` when present.
:::

::: warning
Workbook document modules such as `ThisWorkbook` and sheet modules are protected from `module remove` and `module rename`.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "module new",
  "source": {
    "created": ["src/modules/InvoiceProcessor.bas"],
    "kind": "standard",
    "name": "InvoiceProcessor",
    "path": "src/modules/InvoiceProcessor.bas"
  }
}
```

```json
{
  "status": "ok",
  "command": "module remove",
  "source": {
    "operation": "module.remove",
    "module": "InvoiceAggregator",
    "kind": "standard",
    "removed": ["src/modules/InvoiceAggregator.bas"],
    "requires_push": true
  }
}
```

```json
{
  "status": "ok",
  "command": "module rename",
  "source": {
    "operation": "module.rename",
    "old_name": "OldInvoiceService",
    "new_name": "InvoiceService",
    "kind": "class",
    "renamed": [
      {
        "from": "src/classes/OldInvoiceService.cls",
        "to": "src/classes/InvoiceService.cls"
      }
    ],
    "requires_push": true
  }
}
```

```json
{
  "status": "ok",
  "command": "module install",
  "source": {
    "created": [
      "src/modules/XlflowAssert.bas",
      "src/modules/XlflowRuntime.bas",
      "src/modules/XlflowUI.bas",
      "src/modules/XlflowDebug.bas"
    ]
  }
}
```

## Related

- [init](./init)
- [new](./new)
- [push](./push)
