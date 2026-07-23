# xlflow runner

Manage the persistent xlflow runner marker module used by some workbook execution flows.

## Usage

```bash
xlflow runner install
xlflow runner status
xlflow runner remove
```

## Options and Arguments

| Option / argument | Description                                 | Default |
| ----------------- | ------------------------------------------- | ------- |
| `install`         | Install or update the runner marker module. | -       |
| `status`          | Report whether the marker is present.       | -       |
| `remove`          | Remove the marker module.                   | -       |
| `--json`          | Return structured runner state.             | false   |

## Examples

```bash
xlflow runner status --json
xlflow runner install --json
```

## Notes

::: tip
Most users can rely on `xlflow run`; use `runner` directly when debugging persistent runner state.
:::

`runner` uses the `.NET` bridge on Windows in `auto` mode for `install`, `status`, and `remove`.

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "runner status",
  "runner": { "installed": true, "module": "XlflowRunner" }
}
```

## Related

- [run](./run)

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow runner` when the task matches the command description above. For a goal-oriented workflow, start with the [How-to guides](../guides/) and return here for exact options.

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
