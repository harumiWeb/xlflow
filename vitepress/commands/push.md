# xlflow push

Import edited source files back into the configured workbook.

## Usage

```bash
xlflow push [--backup <always|never>] [--fast] [--changed-only] [--session] [--no-save]
```

## Options and Arguments

| Option / argument | Description                                    | Default                                                          |
| ----------------- | ---------------------------------------------- | ---------------------------------------------------------------- | ------ |
| `--backup <always | never>`                                        | Choose whether to create a backup before modifying the workbook. | always |
| `--fast`          | Use the faster import path when supported.     | false                                                            |
| `--changed-only`  | Import only changed source files.              | false                                                            |
| `--session`       | Push into the managed live workbook session.   | false                                                            |
| `--no-save`       | Leave the session workbook dirty after import. | false                                                            |
| `--json`          | Return import results and warnings.            | false                                                            |

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

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "push",
  "imported": ["src/modules/Main.bas"],
  "backup": "backups/Book-20260517-120000.xlsm",
  "session": { "name": "default", "dirty": true }
}
```

## Related

- [pull](./pull)
- [save](./save)
- [lint](./lint)
