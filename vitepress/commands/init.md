# xlflow init

Initialize an xlflow project from an existing workbook.

## Usage

```bash
xlflow init <workbook> [--with-skill] [--agent <provider>] [--no-update-check]
```

## Options and Arguments

| Option / argument    | Description                                   | Default  |
| -------------------- | --------------------------------------------- | -------- |
| `workbook`           | Existing workbook to bind to the new project. | required |
| `--with-skill`       | Install the bundled AI-agent skill.           | false    |
| `--agent <provider>` | Choose the agent skill provider.              | -        |
| `--no-update-check`  | Skip the startup update check.                | false    |

## Examples

```bash
xlflow init LegacyBook.xlsm
xlflow init LegacyBook.xlsm --with-skill --agent codex --json
```

## Notes

::: important
Imported UserForm projects may use compatibility form code handling. New scaffolds use the safer sidecar layout for form code.
:::

::: tip
Run `xlflow pull --json` after `init` to create the first source-controlled snapshot.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "init",
  "workbook": "LegacyBook.xlsm",
  "config": "xlflow.toml"
}
```

## Related

- [pull](./pull)
- [doctor](./doctor)
- [project structure](../reference/project-structure)
