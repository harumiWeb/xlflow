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

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow list` when the task matches the command description above. For a goal-oriented workflow, start with the [How-to guides](../guides/) and return here for exact options.

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
