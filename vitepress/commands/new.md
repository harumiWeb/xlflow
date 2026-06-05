# xlflow new

Create a new xlflow project and macro-enabled workbook.

## Usage

```bash
xlflow new [workbook] [--with-skill] [--agent <provider>] [--no-update-check]
```

## Options and Arguments

| Option / argument    | Description                                                          | Default   |
| -------------------- | -------------------------------------------------------------------- | --------- |
| `workbook`           | Workbook filename or project name. `.xlsm` is appended when omitted. | Book.xlsm |
| `--with-skill`       | Install the bundled AI-agent skill after scaffolding.                | false     |
| `--agent <provider>` | Select the target agent skill provider, such as `codex`.             | -         |
| `--no-update-check`  | Skip the startup update check.                                       | false     |

## Examples

```bash
xlflow new Sales.xlsm
xlflow new Sales.xlsm --with-skill --agent codex --json
```

## Notes

::: tip
Use `--with-skill --agent codex` when the project will be maintained primarily by an AI coding agent.
:::

::: warning
The generated workbook is macro-enabled. Keep the `.xlsm` extension so Excel can preserve VBA components.
:::

`new` supports explicit `--bridge dotnet` on Windows.

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "new",
  "workbook": "Sales.xlsm",
  "source_root": "src",
  "created": ["xlflow.toml", "src/modules/Main.bas"]
}
```

## Related

- [init](./init)
- [doctor](./doctor)
- [skill](./skill)
