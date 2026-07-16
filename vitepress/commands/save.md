# xlflow save

Save the workbook held by the managed xlflow session.

## Usage

```bash
xlflow save --session
```

## Options and Arguments

| Option / argument | Description                        | Default  |
| ----------------- | ---------------------------------- | -------- |
| `--session`       | Save the managed session workbook. | required |
| `--json`          | Return save status.                | false    |

## Examples

```bash
xlflow save --session
xlflow save --session --json
```

## Notes

> [!IMPORTANT]
> Use this after `push --session --no-save` or `edit --session` when the live workbook should become the persisted workbook.

> [!WARNING]
> When workbook recovery is required, `save` is blocked with
> `workbook_recovery_required`. Do not persist uncertain state. Use managed
> `session stop --discard`, process cleanup, or verified recovery clear.

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "save",
  "session": { "name": "default", "dirty": false }
}
```

## Related

- [session](./session)
- [push](./push)
- [recovery](./recovery)
