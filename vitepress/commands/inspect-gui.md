# xlflow inspect-gui

Scan source for automation-hostile GUI boundaries without opening Excel.

## Usage

```bash
xlflow inspect-gui
```

## Options and Arguments

| Option / argument | Description                                      | Default |
| ----------------- | ------------------------------------------------ | ------- |
| `--json`          | Return detected boundaries and source locations. | false   |

## Examples

```bash
xlflow inspect-gui
xlflow inspect-gui --json
```

## Notes

::: tip
Run `inspect-gui` before headless automation to find likely `MsgBox`, `InputBox`, dialog, or UserForm boundaries.
:::

::: warning
This is static analysis. It is intentionally conservative and does not prove a macro is fully headless-safe.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "inspect-gui",
  "boundaries": [{ "file": "src/modules/Main.bas", "line": 12, "kind": "MsgBox" }]
}
```

## Related

- [lint](./lint)
- [analyze](./analyze)
- [run](./run)
