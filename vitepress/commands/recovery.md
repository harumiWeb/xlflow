# xlflow recovery

Inspect recovery state through `status`, then explicitly clear a workbook
quarantine after Excel or VBA termination could not be confirmed.

## Usage

```bash
xlflow recovery clear
xlflow recovery clear --force
```

## Options

| Option    | Description                                                      | Default |
| --------- | ---------------------------------------------------------------- | ------- |
| `clear`   | Verify that the recorded Excel PID no longer exists, then clear. | -       |
| `--force` | Clear without process verification and emit a safety warning.    | false   |
| `--json`  | Return machine-readable recovery results.                        | false   |

## Why recovery is required

xlflow creates a recovery marker when a command returns but Excel-side
completion is uncertain. Examples include a timed-out macro that may still be
running, a terminated bridge worker, fatal COM/RPC failure, incomplete cleanup,
or a poisoned session whose unsaved state cannot be saved safely.

The recovery marker is not a lock owner. `coordination.busy` still means an
xlflow process currently owns the operating-system workbook lock.
`coordination.recovery_required` means a new workbook operation is unsafe even
when that lock is free.

```bash
xlflow status --json
```

```json
{
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

## Verified clear

Use normal clear only after the affected Excel process has ended:

```bash
xlflow recovery clear --json
```

When no marker exists, the command succeeds idempotently with
`recovery.cleared: false`. When a marker exists, xlflow clears it only if the
marker contains an Excel PID and Windows confirms that PID no longer exists.
A live, unknown, malformed, or unverifiable process state returns
`workbook_recovery_verification_failed` and preserves quarantine.

```json
{
  "status": "ok",
  "command": "recovery clear",
  "recovery": {
    "required": false,
    "cleared": true,
    "forced": false,
    "workbook": "C:\\projects\\sample\\sample.xlsm"
  },
  "logs": []
}
```

## Force clear

```bash
xlflow recovery clear --force --json
```

> [!WARNING]
> Force clear removes only xlflow's marker. It does not terminate VBA, close
> Excel, discard unsaved changes, repair the workbook, or prove that Excel-side
> mutation stopped.

The JSON result includes `recovery.forced: true` and a safety warning. Use this
only after you have manually recovered Excel and accept responsibility for the
remaining state.

## Other recovery paths

- Managed session: `xlflow session stop --discard --json` closes without saving
  and clears recovery only after the owned Excel process is confirmed stopped.
- Known affected PID: `xlflow process cleanup <pid> --json` clears matching
  recovery state only when the process result reports `terminated: true`.
- Complete Excel reset: `xlflow process cleanup --all --yes --json` may clear
  PID-unknown recovery only when a follow-up enumeration confirms no Excel
  process remains.
- External session: close the workbook in Excel without saving, then use
  verified `recovery clear`. xlflow does not automatically close a user-owned
  workbook.

`--wait` and `--wait-timeout` do not wait for or bypass recovery state.

Under WSL, `recovery` delegates to Windows xlflow and uses the same Windows-side
per-user marker as direct Windows commands.

## Related

- [status](./status)
- [session](./session)
- [process](./process)
- [Error Codes](../reference/error-codes)

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow recovery` when the task matches the command description above. For a goal-oriented workflow, start with the [How-to guides](../guides/) and return here for exact options.

## Prerequisites

Check the project configuration and run `xlflow doctor --json` before workbook-backed operations. Source-only commands can run without Excel; commands that read or mutate a workbook require Windows Excel and VBIDE access.

## What this command reads and changes

The command reads the inputs and configuration described in its syntax and examples. Treat source files, the saved workbook, and a live session as separate states; add `--session` when the live workbook is authoritative. Any mutation is reversible only when a backup or explicit session save boundary exists.

## Effect on source-of-truth state

Use `xlflow status --json` before and after the command. A source edit normally requires `push`; a workbook edit normally requires `pull`; a dirty live session requires `save --session` or an intentional discard.

## Common workflows

Combine this command with the relevant [source/workbook/session workflow](../concepts/workbook-session-source), and use `--json` in scripts and agent loops.

## Common failures

Read the structured `error.code`, exit code, and recovery metadata instead of scraping terminal text. The [symptom-oriented troubleshooting guide](../help/troubleshooting) maps installation, execution, session, VS Code, and WSL failures to recovery steps.
