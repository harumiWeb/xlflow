# xlflow init

Initialize an xlflow project from an existing workbook or Excel add-in.

## Usage

```bash
xlflow init <workbook> [--with-module] [--with-skill] [--agent <provider>] [--no-update-check]
```

## Options and Arguments

| Option / argument    | Description                                                      | Default  |
| -------------------- | ---------------------------------------------------------------- | -------- |
| `workbook`           | Existing workbook or add-in to bind to the new project.          | required |
| `--with-module`      | Add bundled helper modules and push them to the copied workbook. | false    |
| `--with-skill`       | Install the bundled AI-agent skill.                              | false    |
| `--agent <provider>` | Choose the agent skill provider.                                 | -        |
| `--no-update-check`  | Skip the startup update check.                                   | false    |

## Examples

```bash
xlflow init LegacyBook.xlsm
xlflow init LegacyAddin.xlam
xlflow init LegacyModel.xlsb
xlflow init LegacyBook.xlsm --with-module
xlflow init LegacyBook.xlsm --with-skill --agent codex --json
```

## Notes

> [!IMPORTANT]
> Imported UserForm projects may use compatibility form code handling. New scaffolds use the safer sidecar layout for form code.

`init` accepts `.xlsm`, `.xlam`, and `.xlsb` files. It copies the input workbook into `build/` without changing its filename, extension, or workbook format. For example, `xlflow init LegacyModel.xlsb` writes `build/LegacyModel.xlsb` and records that path in `xlflow.toml`.

`.xlsb` projects use the normal VBA source layout and Excel COM/VBIDE workflow. Direct OOXML worksheet features such as formula snapshots, workbook cell diff, and pure-Go `pack` are not supported for `.xlsb`.

::: tip
Use `--with-module` when the imported workbook should immediately gain `XlflowAssert`, `XlflowRuntime`, `XlflowUI`, and `XlflowDebug` helpers for workbook-side tests, headless runs, scripted dialogs, and terminal-visible logging.
:::

::: warning
`--with-module` refuses to overwrite existing helper source files. Resolve any `XlflowAssert.bas`, `XlflowRuntime.bas`, `XlflowUI.bas`, or `XlflowDebug.bas` collisions before retrying.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "init",
  "workbook": "build/LegacyAddin.xlam",
  "config": "xlflow.toml"
}
```

## Related

- [pull](./pull)
- [doctor](./doctor)
- [project structure](../reference/project-structure)
