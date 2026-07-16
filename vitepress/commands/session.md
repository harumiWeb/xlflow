# xlflow session

Keep Excel and the configured workbook open across repeated commands.

## Usage

```bash
xlflow session start
xlflow session attach
xlflow session status
xlflow session stop [--discard]
```

## Options and Arguments

| Option / argument | Description                                                        | Default |
| ----------------- | ------------------------------------------------------------------ | ------- |
| `start`           | Open and register the managed workbook session.                    | -       |
| `attach`          | Adopt the already-open configured workbook as an external session. | -       |
| `status`          | Show whether the session is running and dirty.                     | -       |
| `stop`            | Close a managed session, or detach an external session.            | -       |
| `--discard`       | Stop a managed recovery session without saving unsafe state.       | false   |
| `--json`          | Return machine-readable session state.                             | false   |

## Examples

```bash
xlflow session start --json
xlflow session attach --json
xlflow session status --json
xlflow session stop --json
xlflow session stop --discard --json
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

When `coordination.recovery_required` is true, xlflow cannot safely save the
live workbook. Plain `session stop` and `save` are blocked. For a managed
session, use `session stop --discard`; xlflow closes without saving and clears
recovery only after confirming the owned Excel process ended. For an external
session, `--discard` only detaches xlflow metadata and does not close the
user-owned workbook or clear recovery.

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "session",
  "session": { "name": "default", "running": true, "dirty": false },
  "coordination": {
    "busy": false,
    "recovery_required": true,
    "recovery": {
      "reason": "vba_may_still_be_running",
      "operation": "run",
      "recorded_at": "2026-07-16T09:30:00Z",
      "excel_pid": 23456
    }
  }
}
```

`coordination` is a point-in-time observation taken when `session status`
starts. `busy` means the OS workbook lock is currently owned;
`recovery_required` means a persisted quarantine blocks unsafe work even if the
lock is free. Both may briefly be true. A busy workbook may omit owner fields
when current metadata is unavailable. During recovery, status avoids unsafe
workbook COM calls and reports live fields such as `dirty` as unknown,
`source_of_truth: "uncertain"`, and `discard_required: true`. The observation
can change before the response is returned and does not replace the CLI lock
and recovery check performed by later workbook commands. If coordination cannot
be observed, status remains available and returns warning
`coordination_status_unavailable` instead.

## Related

- [push](./push)
- [run](./run)
- [save](./save)
- [recovery](./recovery)
