# Backup and Rollback

The default `push` path is conservative. It creates a rollback-capable workbook backup before replacing VBA components.

```bash
xlflow push --json
```

Backups are stored under `.xlflow/backups/<backup-id>/` and include both the copied workbook file and `metadata.json`.

Automatic backup retention can be enabled in `xlflow.toml`:

```toml
[backup.retention]
enabled = true
max_count = 20
max_age_days = 30
min_keep = 5
max_total_size_mb = 2048
```

It is disabled by default. When enabled, xlflow prunes only backups for the configured workbook after successful backup-producing `push` and `rollback` operations. The newest `min_keep` valid backups are always protected, and invalid or legacy backup directories are skipped rather than deleted. If automatic pruning fails, the successful workbook operation remains successful and the result includes a warning.

List available rollback targets with:

```bash
xlflow backup list --json
```

Restore the newest backup with:

```bash
xlflow rollback --latest --json
```

Or restore a specific backup ID:

```bash
xlflow rollback --backup 20260518-175330-push --json
```

Rollback restores only the workbook file. If source files should match the restored workbook, run:

```bash
xlflow pull --json
```

Fast development loops may use:

```bash
xlflow push --fast --session --no-save --json
```

That skips workbook backup creation for speed and leaves the live session dirty until `xlflow save --session`.

Because no backup is created, automatic retention does not run for this fast path.

If an xlflow session is active for the workbook, `rollback` fails safely instead of replacing the file underneath the live workbook. Stop the session first:

```bash
xlflow session stop --json
```

For review, use `diff` to compare workbook files and optional exported VBA trees:

```bash
xlflow diff before.xlsm after.xlsm --vba-before before-src --vba-after after-src --json
```
