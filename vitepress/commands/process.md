# process

`xlflow process` manages local Excel processes. It is workbook- and configuration-independent: it works on all Excel processes on the local machine, not only those started by `xlflow session start`.

## Usage

```bash
xlflow process list
xlflow process cleanup <pid>
xlflow process cleanup --auto
xlflow process cleanup --all [--yes]
```

## Options

| Flag     | Command         | Description                                                                           |
| -------- | --------------- | ------------------------------------------------------------------------------------- |
| `--auto` | `cleanup`       | Terminate only Excel processes that have no open workbooks.                           |
| `--all`  | `cleanup`       | Force-terminate ALL Excel processes. Prompts for confirmation unless `--yes` is used. |
| `--yes`  | `cleanup --all` | Skip the interactive confirmation prompt for `--all`. Only valid with `--all`.        |

## list

```bash
xlflow process list --json
```

Enumerates all local `EXCEL.EXE` processes. Returns each process's PID, whether
it has open workbooks (`true` = has workbook, `false` = no workbook, `null` =
unknown), and whether a recovery marker identifies that process.

For an affected recovery PID, xlflow does not probe workbook state through COM;
`has_workbook` is `null`, `recovery_required` is `true`, and
`workbook_probe_skipped` is `true`.

If the shared recovery store cannot be enumerated safely or contains malformed
metadata whose affected PID cannot be trusted, `process list` returns
`coordination_recovery_check_failed` before invoking Excel COM.

`cleanup --auto` only targets processes with `has_workbook: false` (confirmed no workbook). Processes with `has_workbook: null` are never targeted by `--auto`.

## cleanup

```bash
xlflow process cleanup 1234 --json
xlflow process cleanup --auto --json
```

Terminates Excel processes in one of three modes:

- `<pid>`: Graceful shutdown of the specified process. Falls back to force-stop if the process persists after a 3-second grace window.
- `--auto`: Graceful shutdown of only those Excel processes that have **no open workbooks**. Workbook-bearing Excel instances are left alone. This is the safe route for cleaning up zombie Excel instances left behind by crashes or failed COM cleanup.
- `--all`: Force-terminates **every** local Excel process, including processes with unsaved workbooks. **Use with extreme caution.**

Successful cleanup also clears matching workbook recovery state:

- `<pid>` clears only markers for that PID when `terminated: true`.
- `--auto` clears only known PIDs that it actually terminates.
- `--all` can clear a marker without a recorded PID only after a follow-up
  enumeration confirms no Excel process remains.

Partial cleanup leaves recovery markers for processes that were not terminated.
When markers are cleared, JSON adds:

```json
{
  "recovery": {
    "cleared": [
      {
        "workbook": "C:\\projects\\sample\\sample.xlsm",
        "excel_pid": 5678
      }
    ],
    "count": 1
  }
}
```

`cleanup --all` always prompts for interactive confirmation:

```text
This will forcibly terminate ALL Excel processes. Unsaved work will be lost. Continue? [y/N]
```

Pass `--yes` to skip the prompt. When `--json` is set, `--yes` is required; otherwise a configuration error is returned.

## Examples

List all Excel processes:

```bash
xlflow process list --json
```

```json
{
  "status": "ok",
  "command": "process list",
  "error": null,
  "process": [
    {
      "pid": 1234,
      "has_workbook": true,
      "workbook_probe_skipped": false,
      "recovery_required": false
    },
    {
      "pid": 5678,
      "has_workbook": false,
      "workbook_probe_skipped": false,
      "recovery_required": false
    },
    {
      "pid": 9012,
      "has_workbook": null,
      "workbook_probe_skipped": true,
      "recovery_required": true
    }
  ],
  "logs": ["found 3 Excel process(es)"]
}
```

Terminate a single Excel process by PID:

```bash
xlflow process cleanup 5678 --json
```

```json
{
  "status": "ok",
  "command": "process cleanup",
  "error": null,
  "process": {
    "action": "cleanup",
    "mode": "pid",
    "total": 1,
    "results": [{ "pid": 5678, "terminated": true, "method": "graceful" }]
  },
  "logs": ["terminated 1 Excel process(es)"]
}
```

Safe cleanup of zombie Excel processes (no open workbooks):

```bash
xlflow process cleanup --auto --json
```

```json
{
  "status": "ok",
  "command": "process cleanup",
  "error": null,
  "process": {
    "action": "cleanup",
    "mode": "auto",
    "total": 1,
    "results": [{ "pid": 5678, "terminated": true, "method": "graceful" }]
  },
  "logs": ["terminated 1 Excel process(es)"]
}
```

Force-terminate all Excel processes with confirmation override:

```bash
xlflow process cleanup --all --yes --json
```

```json
{
  "status": "ok",
  "command": "process cleanup",
  "error": null,
  "process": {
    "action": "cleanup",
    "mode": "all",
    "total": 2,
    "results": [
      { "pid": 1234, "terminated": true, "method": "force" },
      { "pid": 5678, "terminated": true, "method": "force" }
    ]
  },
  "logs": ["terminated 2 Excel process(es)"]
}
```

When force-stop fails for an individual process, `method` is `"none"` and the command returns `status: "failed"` with error code `process_termination_failed`. For example:

```json
{
  "status": "failed",
  "command": "process cleanup",
  "error": {
    "code": "process_termination_failed",
    "message": "1 of 3 Excel process(es) failed to terminate"
  },
  "process": {
    "action": "cleanup",
    "mode": "all",
    "total": 3,
    "results": [
      { "pid": 1234, "terminated": true, "method": "force" },
      { "pid": 5678, "terminated": true, "method": "force" },
      { "pid": 9012, "terminated": false, "method": "none" }
    ]
  },
  "logs": null
}
```

## Danger of cleanup --all

`cleanup --all` is a **destructive operation**. It forcibly terminates every Excel process on the machine regardless of whether unsaved workbooks are open. Any workbook changes not yet saved to disk will be lost.

Use `--auto` when you only want to clean up orphaned Excel processes with no open workbooks. Reserve `--all` for situations where you need a complete Excel restart and understand the data-loss implications.

## Related

- [recovery](./recovery)
- [session](./session)
- [status](./status)
