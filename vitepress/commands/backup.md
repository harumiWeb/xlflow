# xlflow backup

List rollback-capable workbook backups managed by xlflow.

## Usage

```bash
xlflow backup list
xlflow backup prune [--keep-last <n>] [--older-than <duration>] [--max-total-size <size>] [--dry-run] [--all-workbooks] [--include-invalid] [--include-legacy]
```

## Options and Arguments

| Option / argument   | Description                                              | Default  |
| ------------------- | -------------------------------------------------------- | -------- |
| `list`              | List workbook-file backups for the configured workbook.  | required |
| `prune`             | Preview or delete backups matching an explicit policy.   | optional |
| `--keep-last <n>`   | Protect the newest `n` valid backups.                    | none     |
| `--older-than`      | Delete valid backups older than a duration like `30d`.   | none     |
| `--max-total-size`  | Delete oldest backups until total size is within limit.  | none     |
| `--dry-run`         | Preview candidates without deleting files.               | false    |
| `--all-workbooks`   | Evaluate valid backups for all managed workbooks.        | false    |
| `--include-invalid` | Include invalid managed backup directories for cleanup.  | false    |
| `--include-legacy`  | Include legacy directories without metadata for cleanup. | false    |
| `--json`            | Return backup metadata for scripts and agent workflows.  | false    |

## Examples

```bash
xlflow backup list
xlflow backup list --json
xlflow backup prune --keep-last 20 --dry-run --json
xlflow backup prune --keep-last 20 --older-than 30d --max-total-size 2GB --json
```

## Notes

> [!IMPORTANT]
> `backup list` only shows metadata-backed workbook backups that xlflow can restore with `rollback`. Older component-export backup directories are ignored.

Manual `backup prune` is independent of `[backup.retention]`; it does not require automatic retention to be enabled. By default it is scoped to the configured workbook. Invalid entries and legacy directories are skipped unless explicitly included with cleanup flags. Backups preserve the configured workbook extension and file format, including `.xlsb`.

Automatic pruning is configured separately in `xlflow.toml` under `[backup.retention]`. It is disabled by default and runs only after successful backup-producing `push` or `rollback` operations.

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "backup list",
  "backups": [
    {
      "id": "20260518-175330-push",
      "created_at": "2026-05-18T17:53:31+09:00",
      "reason": "before-push",
      "workbook": "build/Book.xlsm",
      "path": ".xlflow/backups/20260518-175330-push/Book.xlsm"
    }
  ]
}
```

## Related

- [rollback](./rollback)
- [push](./push)
- [Backup and Rollback](../concepts/backup-and-rollback)

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow backup` when the task matches the command description above. For a goal-oriented workflow, start with the [How-to guides](../guides/) and return here for exact options.

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
