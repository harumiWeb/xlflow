# xlflow backup

List rollback-capable workbook backups managed by xlflow.

## Usage

```bash
xlflow backup list
```

## Options and Arguments

| Option / argument | Description                                             | Default  |
| ----------------- | ------------------------------------------------------- | -------- |
| `list`            | List workbook-file backups for the configured workbook. | required |
| `--json`          | Return backup metadata for scripts and agent workflows. | false    |

## Examples

```bash
xlflow backup list
xlflow backup list --json
```

## Notes

> [!IMPORTANT]
> `backup list` only shows metadata-backed workbook backups that xlflow can restore with `rollback`. Older component-export backup directories are ignored.

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
