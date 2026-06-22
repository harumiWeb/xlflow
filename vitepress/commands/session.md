# xlflow session

Keep Excel and the configured workbook open across repeated commands.

## Usage

```bash
xlflow session start
xlflow session status
xlflow session stop
```

## Options and Arguments

| Option / argument | Description                                     | Default |
| ----------------- | ----------------------------------------------- | ------- |
| `start`           | Open and register the managed workbook session. | -       |
| `status`          | Show whether the session is running and dirty.  | -       |
| `stop`            | Close or detach the managed session.            | -       |
| `--json`          | Return machine-readable session state.          | false   |

## Examples

```bash
xlflow session start --json
xlflow session status --json
xlflow session stop --json
```

## Notes

::: tip
Use sessions for fast AI-agent loops: `push --session --no-save`, `run --session`, inspect results, then `save --session`.
:::

::: warning
A dirty session may report `save_required`. That warning means disk does not yet contain the live workbook changes.
:::

`session` uses the `.NET` bridge on Windows in `auto` mode for `start`, `status`, `save`, and `stop`. Deprecated `--bridge powershell` remains explicit opt-in for v0.15.0 only and emits a removal warning.

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "session status",
  "session": { "name": "default", "running": true, "dirty": false }
}
```

## Related

- [push](./push)
- [run](./run)
- [save](./save)
