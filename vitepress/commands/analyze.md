# xlflow analyze

Analyze VBA source for runtime-risk patterns without Excel COM.

## Usage

```bash
xlflow analyze
```

## Options and Arguments

| Option / argument | Description                          | Default |
| ----------------- | ------------------------------------ | ------- |
| `--json`          | Return structured analysis findings. | false   |

## Examples

```bash
xlflow analyze
xlflow analyze --json
```

## Notes

::: tip
Use `analyze` for fast source-level feedback before opening Excel.
:::

::: important
Findings that block automation return a failure status and exit code `1`.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "error",
  "command": "analyze",
  "findings": [{ "file": "src/modules/Main.bas", "line": 20, "code": "interactive_input" }]
}
```

## Related

- [lint](./lint)
- [check](./check)
- [inspect-gui](./inspect-gui)
