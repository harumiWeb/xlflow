# xlflow new

Create a new xlflow project and macro-enabled workbook or Excel add-in.

## Usage

```bash
xlflow new [workbook] [--with-skill] [--agent <provider>] [--no-update-check]
```

## Options and Arguments

| Option / argument    | Description                                                                                                               | Default   |
| -------------------- | ------------------------------------------------------------------------------------------------------------------------- | --------- |
| `workbook`           | Workbook filename or project name. `.xlsm` is appended when omitted. Explicit `.xlsm`, `.xlam`, and `.xlsb` are accepted. | Book.xlsm |
| `--with-skill`       | Install the bundled AI-agent skill after scaffolding.                                                                     | false     |
| `--agent <provider>` | Select the target agent skill provider, such as `codex`.                                                                  | -         |
| `--no-update-check`  | Skip the startup update check.                                                                                            | false     |

## Examples

```bash
xlflow new Sales
xlflow new Sales.xlsm
xlflow new SalesAddin.xlam
xlflow new LargeModel.xlsb
xlflow new Sales.xlsm --with-skill --agent codex --json
```

## Notes

::: tip
Use `--with-skill --agent codex` when the project will be maintained primarily by an AI coding agent.
:::

::: warning
The generated workbook must use `.xlsm`, `.xlam`, or `.xlsb` so Excel can preserve VBA components. Omit the extension only when you want the default `.xlsm` project format.
:::

For `.xlam` projects, `new` creates the add-in file in the project's `build/` directory only. It does not install or register the add-in in Excel.

For `.xlsb` projects, xlflow uses Excel COM/VBIDE for VBA synchronization and creates the workbook with Excel file format `50`. Direct OOXML worksheet features such as formula snapshots, workbook cell diff, and pure-Go `pack` do not support `.xlsb`.

`new` uses the `.NET` bridge on Windows in `auto` mode.

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "new",
  "workbook": "build/SalesAddin.xlam",
  "source_root": "src",
  "created": ["xlflow.toml", "src/modules/Main.bas"]
}
```

## Related

- [init](./init)
- [doctor](./doctor)
- [skill](./skill)

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow new` when the task matches the command description above. For a goal-oriented workflow, start with the [How-to guides](../guides/) and return here for exact options.

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
