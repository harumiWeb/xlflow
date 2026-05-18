# xlflow pull

Export workbook VBA components and form artifacts into configured source directories.

## Usage

```bash
xlflow pull [--session] [--keepalive] [--keepalive-interval <duration>]
```

## Options and Arguments

| Option / argument | Description                                  | Default |
| ----------------- | -------------------------------------------- | ------- |
| `--session`       | Pull from the managed live session workbook. | false   |
| `--keepalive`     | Reuse the bridge process.                    | false   |
| `--json`          | Report exported files and warnings.          | false   |

## Examples

```bash
xlflow pull
xlflow pull --session --json
```

## Notes

::: tip
Pull before editing if the workbook may contain newer VBA than the source tree.
:::

> [!IMPORTANT]
> UserForm Designer state and code-behind may be written to separate sidecar paths depending on project configuration.

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "pull",
  "workbook": "Book.xlsm",
  "written": ["src/modules/Main.bas", "src/forms/specs/UserForm1.yaml"]
}
```

## Related

- [push](./push)
- [diff](./diff)
- [project structure](../reference/project-structure)
