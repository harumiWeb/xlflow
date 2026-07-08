# xlflow session

Keep Excel and the configured workbook open across repeated commands.

## Usage

```bash
xlflow session start
xlflow session attach
xlflow session status
xlflow session stop
```

## Options and Arguments

| Option / argument | Description                                                        | Default |
| ----------------- | ------------------------------------------------------------------ | ------- |
| `start`           | Open and register the managed workbook session.                    | -       |
| `attach`          | Adopt the already-open configured workbook as an external session. | -       |
| `status`          | Show whether the session is running and dirty.                     | -       |
| `stop`            | Close a managed session, or detach an external session.            | -       |
| `--json`          | Return machine-readable session state.                             | false   |

## Examples

```bash
xlflow session start --json
xlflow session attach --json
xlflow session status --json
xlflow session stop --json
```

## Notes

::: tip
Use sessions for fast AI-agent loops: `push --session --no-save`, `run --session`, inspect results, then `save --session`.
:::

::: tip
When the workbook is already open in Excel, use `xlflow session attach` instead of `session start` so xlflow commands operate on the workbook you are viewing.
:::

::: warning
A dirty session may report `save_required`. That warning means disk does not yet contain the live workbook changes.
:::

`session start` creates a managed session owned by xlflow. `session attach` creates an external session for a workbook that was opened by a user. `session stop` closes and quits managed sessions, but only detaches external sessions; it does not close Excel.

`session` uses the `.NET` bridge on Windows in `auto` mode for `start`, `attach`, `status`, `save`, and `stop`.

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
