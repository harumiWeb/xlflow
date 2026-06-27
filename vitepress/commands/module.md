# xlflow module

Create VBA module source files or install bundled xlflow helper modules into an existing project.

## Usage

```bash
xlflow module new <name> --type standard
xlflow module new <name> --type class
xlflow module install [--push]
```

## Options and Arguments

| Option / argument | Description                                                                                | Default  |
| ----------------- | ------------------------------------------------------------------------------------------ | -------- |
| `new <name>`      | Create a new standard or class module source file.                                         | -        |
| `--type`          | Required for `new`; must be `standard` or `class`.                                         | required |
| `install`         | Install the bundled helper modules into the configured module source root.                 | required |
| `--push`          | Push the installed helper modules into the configured workbook after writing source files. | false    |

## Examples

```bash
xlflow module new InvoiceProcessor --type standard
xlflow module new InvoiceService --type class --json
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
