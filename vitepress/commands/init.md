# xlflow init

Initialize an xlflow project from an existing workbook or Excel add-in.

## Usage

```bash
xlflow init <workbook> [--userform-code-source frm|sidecar] [--with-module] [--with-skill] [--agent <provider>] [--no-update-check]
```

## Options and Arguments

| Option / argument               | Description                                                                              | Default   |
| ------------------------------- | ---------------------------------------------------------------------------------------- | --------- |
| `workbook`                      | Existing workbook or add-in to bind to the new project.                                  | required  |
| `--userform-code-source <mode>` | UserForm code source for imported projects: `sidecar` or explicit `frm` artifact import. | `sidecar` |
| `--with-module`                 | Add bundled helper modules and push them to the copied workbook.                         | false     |
| `--with-skill`                  | Install the bundled AI-agent skill.                                                      | false     |
| `--agent <provider>`            | Choose the agent skill provider.                                                         | -         |
| `--no-update-check`             | Skip the startup update check.                                                           | false     |

## Examples

```bash
xlflow init LegacyBook.xlsm
xlflow init LegacyAddin.xlam
xlflow init LegacyModel.xlsb
xlflow init LegacyBook.xlsm --userform-code-source frm
xlflow init LegacyBook.xlsm --with-module
xlflow init LegacyBook.xlsm --with-skill --agent codex --json
```

## Notes

> [!IMPORTANT]
> Imported UserForm projects default to `sidecar`: the bootstrap pull creates canonical Designer specs and sidecar code. Pass `--userform-code-source frm` explicitly when importing `.frm` artifacts as the code source, including workflows that rely on custom-control artifact fidelity.

`init` accepts `.xlsm`, `.xlam`, and `.xlsb` files. It copies the input workbook into `build/` without changing its filename, extension, or workbook format. For example, `xlflow init LegacyModel.xlsb` writes `build/LegacyModel.xlsb` and records that path in `xlflow.toml`.

Any source-preflight step reached during initialization uses the generated
project's `[preflight].allowed_diagnostics` policy. Registry diagnostics may be
waived explicitly, but unreadable source, duplicate components, and UserForm
artifact integrity failures remain blocking.

`.xlsb` projects use the normal VBA source layout and Excel COM/VBIDE workflow. Direct OOXML worksheet features such as formula snapshots, workbook cell diff, and pure-Go `pack` are not supported for `.xlsb`.

`init` defaults to `[userform].code_source = "sidecar"`, runs the bootstrap `pull`, and creates Designer specs under `src/forms/specs/*.yaml` plus sidecar code under `src/forms/code/*.bas` for imported UserForms. With explicit `--userform-code-source frm`, embedded code from imported `.frm` files remains authoritative. `.frm` and `.frx` are compatibility/import artifacts and generated outputs; canonical Designer specs do not require them as build inputs.

::: tip
Use `--with-module` when the imported workbook should immediately gain `XlflowAssert`, `XlflowRuntime`, `XlflowUI`, and `XlflowDebug` helpers for workbook-side tests, headless runs, scripted dialogs, and terminal-visible logging.
:::

::: warning
`--with-module` writes helpers under `src/modules/Xlflow/` and refuses to overwrite either those target files or legacy root-level helper files. Resolve any `XlflowAssert.bas`, `XlflowRuntime.bas`, `XlflowUI.bas`, or `XlflowDebug.bas` collision before retrying.
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

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow init` when the task matches the command description above. For a goal-oriented workflow, start with the [How-to guides](../guides/) and return here for exact options.

## Prerequisites

Check the project configuration and run `xlflow doctor --json` before workbook-backed operations. Source-only commands can run without Excel; commands that read or mutate a workbook require Windows Excel and VBIDE access.

## What this command reads and changes

The command reads the inputs and configuration described in its syntax and examples. Treat source files, the saved workbook, and a live session as separate states; add `--session` when the live workbook is authoritative. Any mutation is reversible only when a backup or explicit session save boundary exists.

## Effect on source-of-truth state

Use `xlflow status --json` before and after the command. A source edit normally requires `push`; a workbook edit normally requires `pull`; a dirty live session requires `save --session` or an intentional discard.

## Common workflows

Combine this command with the relevant [source/workbook/session workflow](../concepts/workbook-session-source), and use `--json` in scripts and agent loops.

## Common failures

Read the structured `error.code`, exit code, and recovery metadata instead of scraping terminal text. The [symptom-oriented troubleshooting guide](../help/troubleshooting) maps installation, execution, session, VS Code, and WSL failures to recovery steps.
