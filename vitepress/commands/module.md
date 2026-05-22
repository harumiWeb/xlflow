# xlflow module

Install the bundled xlflow helper modules into an existing project.

## Usage

```bash
xlflow module install [--push]
```

## Options and Arguments

| Option / argument | Description                                                                                | Default  |
| ----------------- | ------------------------------------------------------------------------------------------ | -------- |
| `install`         | Install the bundled helper modules into the configured module source root.                 | required |
| `--push`          | Push the installed helper modules into the configured workbook after writing source files. | false    |

## Examples

```bash
xlflow module install
xlflow module install --push --json
```

## Notes

::: tip
`module install` writes `XlflowAssert.bas`, `XlflowRuntime.bas`, `XlflowUI.bas`, and `XlflowDebug.bas` into the configured `[src].modules` root.
:::

::: warning
The command refuses to overwrite existing helper source files. Remove or rename colliding files before retrying.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

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
