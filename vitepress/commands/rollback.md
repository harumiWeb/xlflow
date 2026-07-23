# xlflow rollback

Restore the configured workbook from an xlflow-managed workbook backup.

## Usage

```bash
xlflow rollback --latest
xlflow rollback --backup <backup-id>
```

## Options and Arguments

| Option / argument      | Description                                            | Default |
| ---------------------- | ------------------------------------------------------ | ------- |
| `--latest`             | Restore the newest backup for the configured workbook. | false   |
| `--backup <backup-id>` | Restore a specific backup ID.                          | none    |
| `--json`               | Return rollback metadata, warnings, and hints.         | false   |

## Examples

```bash
xlflow rollback --latest --json
xlflow rollback --backup 20260518-175330-push --json
```

## Notes

> [!IMPORTANT]
> Rollback restores only the workbook file. It does not rewrite `src/` automatically.

Rollback restores the backed-up file as-is, preserving the workbook extension and file format, including `.xlsb`.

::: warning
If the workbook is attached to an active xlflow session, rollback fails with `workbook_in_use`. Stop the session first, then rerun rollback.
:::

If `[backup.retention].enabled = true`, rollback automatically prunes old backups only after the pre-rollback safety backup is created and the selected backup is restored successfully. The safety backup remains in the rollback result and participates in retention evaluation. Automatic pruning failures are warnings and do not fail the successful rollback.

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "rollback",
  "rollback": {
    "restored_from": {
      "id": "20260518-175330-push",
      "path": ".xlflow/backups/20260518-175330-push/Book.xlsm",
      "reason": "before-push",
      "created_at": "2026-05-18T17:53:31+09:00"
    },
    "safety_backup": {
      "id": "20260518-175431-pre-rollback",
      "path": ".xlflow/backups/20260518-175431-pre-rollback/Book.xlsm"
    },
    "target": {
      "path": "build/Book.xlsm"
    }
  },
  "warnings": [
    {
      "code": "source_out_of_sync",
      "message": "Rollback restored only the workbook file. Source files under `src/` were not changed and may now be out of sync."
    }
  ],
  "hints": [
    {
      "code": "verify_workbook",
      "message": "Run `xlflow inspect --json` to verify the restored workbook state."
    },
    {
      "code": "sync_source",
      "message": "Run `xlflow pull --json` if you want source files to match the restored workbook."
    }
  ]
}
```

## Related

- [backup](./backup)
- [pull](./pull)
- [session](./session)
- [Backup and Rollback](../concepts/backup-and-rollback)

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow rollback` when the task matches the command description above. For a goal-oriented workflow, start with the [How-to guides](../guides/) and return here for exact options.

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
