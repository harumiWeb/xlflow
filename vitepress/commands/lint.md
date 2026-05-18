# xlflow lint

Lint VBA source files for agent-hostile and compile-dialog-prone patterns.

## Usage

```bash
xlflow lint
```

## Options and Arguments

| Option / argument | Description                    | Default |
| ----------------- | ------------------------------ | ------- |
| `--json`          | Return structured lint issues. | false   |

## Examples

```bash
xlflow lint
xlflow lint --json
```

## Notes

> [!IMPORTANT]
> Syntax-safety checks are always enabled for patterns that could surface as modal VBE compile dialogs.

::: tip
Use `lint --json` in agent loops before `push` to catch source problems while Excel is still closed.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "error",
  "command": "lint",
  "issues": [
    {
      "file": "src/modules/Main.bas",
      "line": 7,
      "code": "vba_syntax_risk",
      "severity": "error"
    }
  ]
}
```

## Related

- [analyze](./analyze)
- [check](./check)
- [error codes](../reference/error-codes)
