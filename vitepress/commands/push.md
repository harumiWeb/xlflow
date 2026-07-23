# xlflow push

Import edited source files back into the configured workbook.

## Usage

```bash
xlflow push [--backup <always|never>] [--fast] [--changed-only] [--session] [--no-save]
```

## Options and Arguments

| Option / argument          | Description                                                                                | Default |
| -------------------------- | ------------------------------------------------------------------------------------------ | ------- |
| `--backup <always\|never>` | Choose whether to create a rollback-capable workbook backup before modifying the workbook. | always  |
| `--fast`                   | Use the faster import path when supported.                                                 | false   |
| `--changed-only`           | Import only changed source files.                                                          | false   |
| `--session`                | Push into the managed live workbook session.                                               | false   |
| `--no-save`                | Leave the session workbook dirty after import.                                             | false   |
| `--json`                   | Return import results and warnings.                                                        | false   |

## Examples

```bash
xlflow push --backup always --json
xlflow push --session --fast --no-save --json
```

## Notes

> [!IMPORTANT]
> `push` runs source preflight before opening Excel so modal compile dialogs are caught as structured CLI errors whenever possible.

When `[vba.line_numbers].enabled = true`, `push` updates folder annotations and then adds temporary physical-line labels to its prepared import copies so VBA `Erl` reports useful locations. The tracked source is not changed. Labels use fixed-width space padding and no colon; no `push` flag is provided for this feature. xlflow stops safely instead of instrumenting code that contains existing or mismatched numeric labels, or numeric `GoTo`, `GoSub`, or `Resume` targets.

::: warning
`--session --no-save` leaves the live workbook newer than disk. Run `xlflow save --session` when the changes should persist.
:::

::: tip
The default backup is a workbook-file snapshot under `.xlflow/backups/<backup-id>/` with `metadata.json`. Use `xlflow backup list --json` to inspect rollback targets.
:::

If `[backup.retention].enabled = true`, `push` automatically prunes old backups after a successful workbook update only when a new backup was created. It does not run after `--backup never`, `--fast`, unchanged `--changed-only` no-op pushes, or failed pushes. Automatic pruning is scoped to the configured workbook, skips invalid and legacy entries, and reports pruning failures as warnings without failing the successful push.

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "push",
  "backup": {
    "id": "20260518-175330-push",
    "mode": "always",
    "path": ".xlflow/backups/20260518-175330-push/Book.xlsm",
    "reason": "before-push"
  },
  "source": {
    "changed": true,
    "changed_only": false,
    "state": ".xlflow/state/push.json"
  },
  "workbook": {
    "path": "build/Book.xlsm",
    "saved": true,
    "session": false
  }
}
```

## Related

- [backup](./backup)
- [rollback](./rollback)
- [pull](./pull)
- [save](./save)
- [lint](./lint)

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow push` when the task matches the command description above. For a goal-oriented workflow, start with the [How-to guides](../guides/) and return here for exact options.

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
