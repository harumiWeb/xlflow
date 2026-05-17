# xlflow list

List workbook resources. The public resource command is `list forms`.

## Usage

```bash
xlflow list forms [--session] [--keepalive] [--keepalive-interval <duration>]
```

## Options and Arguments

| Option / argument | Description                                  | Default  |
| ----------------- | -------------------------------------------- | -------- |
| `forms`           | List UserForms and expected source paths.    | required |
| `--session`       | Read from the managed live workbook session. | false    |
| `--json`          | Return form metadata for scripts.            | false    |

## Examples

```bash
xlflow list forms
xlflow list forms --session --json
```

## Notes

::: tip
Use `list forms` before `form snapshot`, `form build`, or `form export-image` to confirm names.
:::

::: important
Listing forms is metadata-oriented; it does not execute UserForm runtime code.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "list forms",
  "forms": [{ "name": "UserForm1", "designer": "src/forms/specs/UserForm1.yaml" }]
}
```

## Related

- [form](./form)
- [inspect](./inspect)
