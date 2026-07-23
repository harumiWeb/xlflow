# xlflow inspect-gui

Scan source for automation-hostile GUI boundaries without opening Excel.

## Usage

```bash
xlflow inspect-gui
```

## Options and Arguments

| Option / argument | Description                                      | Default |
| ----------------- | ------------------------------------------------ | ------- |
| `--json`          | Return detected boundaries and source locations. | false   |

## Examples

```bash
xlflow inspect-gui
xlflow inspect-gui --json
```

## Notes

::: tip
Run `inspect-gui` before headless automation to find likely `MsgBox`, `InputBox`, dialog, or UserForm boundaries.
:::

::: warning
This is static analysis. It is intentionally conservative and does not prove a macro is fully headless-safe.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "inspect-gui",
  "boundaries": [{ "file": "src/modules/Main.bas", "line": 12, "kind": "MsgBox" }]
}
```

## Related

- [lint](./lint)
- [analyze](./analyze)
- [run](./run)

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow inspect gui` when the task matches the command description above. For a goal-oriented workflow, start with the [How-to guides](../guides/) and return here for exact options.

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
