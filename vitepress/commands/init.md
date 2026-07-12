# xlflow init

Initialize an xlflow project from an existing workbook or Excel add-in.

## Usage

```bash
xlflow init <workbook> [--userform-code-source frm|sidecar] [--with-module] [--with-skill] [--agent <provider>] [--no-update-check]
```

## Options and Arguments

| Option / argument               | Description                                                      | Default  |
| ------------------------------- | ---------------------------------------------------------------- | -------- |
| `workbook`                      | Existing workbook or add-in to bind to the new project.          | required |
| `--userform-code-source <mode>` | UserForm code source for imported projects: `frm` or `sidecar`.  | `frm`    |
| `--with-module`                 | Add bundled helper modules and push them to the copied workbook. | false    |
| `--with-skill`                  | Install the bundled AI-agent skill.                              | false    |
| `--agent <provider>`            | Choose the agent skill provider.                                 | -        |
| `--no-update-check`             | Skip the startup update check.                                   | false    |

## Examples

```bash
xlflow init LegacyBook.xlsm
xlflow init LegacyAddin.xlam
xlflow init LegacyModel.xlsb
xlflow init LegacyBook.xlsm --userform-code-source sidecar
xlflow init LegacyBook.xlsm --with-module
xlflow init LegacyBook.xlsm --with-skill --agent codex --json
```

## Notes

> [!IMPORTANT]
> Imported UserForm projects default to compatibility `frm` code handling so existing code-behind embedded in exported `.frm` files remains authoritative. Use `--userform-code-source sidecar` to import directly into the modern sidecar layout.

`init` accepts `.xlsm`, `.xlam`, and `.xlsb` files. It copies the input workbook into `build/` without changing its filename, extension, or workbook format. For example, `xlflow init LegacyModel.xlsb` writes `build/LegacyModel.xlsb` and records that path in `xlflow.toml`.

`.xlsb` projects use the normal VBA source layout and Excel COM/VBIDE workflow. Direct OOXML worksheet features such as formula snapshots, workbook cell diff, and pure-Go `pack` are not supported for `.xlsb`.

When `--userform-code-source sidecar` is selected, `init` writes `[userform].code_source = "sidecar"`, runs the bootstrap `pull`, and creates Designer specs under `src/forms/specs/*.yaml` for imported UserForms. Sidecar code lives under `src/forms/code/*.bas`; `.frm` and `.frx` remain generated artifacts used by pull, push, and build workflows.

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
