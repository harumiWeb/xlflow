# xlflow push

Import edited source files back into the configured workbook.

## Usage

```bash
xlflow push [--backup <always|never>] [--fast] [--changed-only] [--session] [--no-save]
```

## Options and Arguments

| Option / argument          | Description                                                                                | Default |
| -------------------------- | ------------------------------------------------------------------------------------------ | ------- |
| `--backup <always\|never>` | Choose whether to create a rollback-capable workbook backup before modifying the workbook. | always  |
| `--fast`                   | Use the faster import path when supported.                                                 | false   |
| `--changed-only`           | Import only changed source files.                                                          | false   |
| `--session`                | Push into the managed live workbook session.                                               | false   |
| `--no-save`                | Leave the session workbook dirty after import.                                             | false   |
| `--json`                   | Return import results and warnings.                                                        | false   |

## Examples

```bash
xlflow push --backup always --json
xlflow push --session --fast --no-save --json
```

## Notes

> [!IMPORTANT]
> `push` runs source preflight before opening Excel so modal compile dialogs are caught as structured CLI errors whenever possible.

::: warning
`--session --no-save` leaves the live workbook newer than disk. Run `xlflow save --session` when the changes should persist.
:::

::: tip
The default backup is a workbook-file snapshot under `.xlflow/backups/<backup-id>/` with `metadata.json`. Use `xlflow backup list --json` to inspect rollback targets.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "push",
  "backup": {
    "id": "20260518-175330-push",
    "mode": "always",
    "path": ".xlflow/backups/20260518-175330-push/Book.xlsm",
    "reason": "before-push"
  },
  "source": {
    "changed": true,
    "changed_only": false,
    "state": ".xlflow/state/push.json"
  },
  "workbook": {
    "path": "build/Book.xlsm",
    "saved": true,
    "session": false
  }
}
```

## Related

- [backup](./backup)
- [rollback](./rollback)
- [pull](./pull)
- [save](./save)
- [lint](./lint)
