# xlflow list

List workbook resources. The public resource command is `list forms`.

## Usage

```bash
xlflow [--wait] list forms [--session]
```

## Options and Arguments

| Option / argument | Description                                  | Default  |
| ----------------- | -------------------------------------------- | -------- |
| `forms`           | List UserForms and expected source paths.    | required |
| `--session`       | Read from the managed live workbook session. | false    |
| `--json`          | Return form metadata for scripts.            | false    |
| `--wait`          | Wait up to 30 seconds for the workbook lock. | false    |

## Examples

```bash
xlflow list forms
xlflow list forms --session --json
xlflow --wait --wait-timeout 15s list forms --json
```

## Notes

::: tip
Use `list forms` before `form snapshot`, `form build`, or `form export-image` to confirm names.
:::

> [!IMPORTANT]
> Listing forms is metadata-oriented; it does not execute UserForm runtime code.

`list forms` uses the `.NET` bridge on Windows in `auto` mode.
It shares the configured workbook lock with Designer, execution, pull, and push
operations. Contention returns `workbook_busy`; global `--wait` performs an
explicit bounded acquisition wait without retrying the list handler itself.

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
