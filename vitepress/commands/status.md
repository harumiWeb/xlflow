# xlflow status

Show project, source, workbook, and session state in one read-only command.

## Usage

```bash
xlflow status
xlflow status --json
```

## Options and Arguments

| Option / argument | Description                    | Default |
| ----------------- | ------------------------------ | ------- |
| `--json`          | Return machine-readable state. | false   |

## Examples

```bash
xlflow status --json
```

## Notes

::: tip
Use `status` before editing to confirm whether the saved workbook, live session, and source files are in sync.
:::

::: tip
AI agents should prefer `xlflow status --json` before pushing or running to avoid working against stale workbook state.
:::

::: warning
`src_newer_than_workbook` is a heuristic based on file modification times. Clock skew or manual copies can cause false results.
:::

`status` also distinguishes current lock ownership from recovery quarantine.
`busy: true` means another xlflow operation owns the operating-system workbook
lock. `recovery_required: true` means Excel-side completion is uncertain and
unsafe workbook commands remain blocked even when `busy` is false.

## JSON Output Example

```json
{
  "status": "ok",
  "command": "status",
  "project": {
    "root": ".",
    "workbook_path": "build/Book.xlsm",
    "src_paths": ["src/modules", "src/classes", "src/forms", "src/workbook"],
    "project_name": "sample"
  },
  "session": {
    "active": false,
    "workbook_path": "build/Book.xlsm",
    "workbook_name": "Book.xlsm",
    "dirty": false,
    "running": false,
    "workbook_open": false,
    "metadata": null,
    "save_required": false,
    "live_newer_than_disk": false,
    "source_of_truth": "saved_workbook"
  },
  "coordination": {
    "busy": false,
    "recovery_required": false
  },
  "state": {
    "src_newer_than_workbook": false,
    "live_session_newer_than_disk": false,
    "workbook_saved": true,
    "source_of_truth": "saved_workbook",
    "workbook_last_modified_at": "2026-05-23T10:00:00Z",
    "latest_source_modified_at": "2026-05-22T10:00:00Z"
  },
  "warnings": [],
  "hints": [],
  "error": null,
  "logs": ["status reported"]
}
```

While recovery is required, `status` does not call unsafe workbook COM APIs. It
reports `session.dirty: null`, `session.source_of_truth: "uncertain"`,
`session.discard_required: true`, and nested recovery details under
`coordination.recovery`. Use [recovery](./recovery) to choose a safe clearing
path.

If recovery metadata cannot be read safely, `status` still does not report a
clean workbook. It sets `coordination.recovery_check_failed: true`, keeps
`busy` / `recovery_required` unknown (`null`), and reports the session/source of
truth as uncertain.

## Related

- [session](./session)
- [inspect](./inspect)
- [push](./push)
- [save](./save)
- [recovery](./recovery)
