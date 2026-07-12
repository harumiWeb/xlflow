# xlflow backup

List rollback-capable workbook backups managed by xlflow.

## Usage

```bash
xlflow backup list
xlflow backup prune [--keep-last <n>] [--older-than <duration>] [--max-total-size <size>] [--dry-run]
```

## Options and Arguments

| Option / argument  | Description                                             | Default  |
| ------------------ | ------------------------------------------------------- | -------- |
| `list`             | List workbook-file backups for the configured workbook. | required |
| `prune`            | Preview or delete backups matching an explicit policy.  | optional |
| `--keep-last <n>`  | Protect the newest `n` valid backups.                   | none     |
| `--older-than`     | Delete valid backups older than a duration like `30d`.  | none     |
| `--max-total-size` | Delete oldest backups until total size is within limit. | none     |
| `--dry-run`        | Preview candidates without deleting files.              | false    |
| `--json`           | Return backup metadata for scripts and agent workflows. | false    |

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

Manual `backup prune` is independent of `[backup.retention]`; it does not require automatic retention to be enabled. By default it is scoped to the configured workbook. Invalid entries and legacy directories are skipped unless explicitly included with cleanup flags.

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
