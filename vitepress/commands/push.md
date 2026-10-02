# xlflow push

Import edited source files back into the configured workbook.

## Usage

```bash
xlflow push [--backend <excel|file>] [--backup <always|never>] [--fast] [--changed-only] [--session] [--no-save]
```

## Options and Arguments

| Option / argument          | Description                                                                                | Default |
| -------------------------- | ------------------------------------------------------------------------------------------ | ------- |
| `--backend <excel\|file>`  | Select the Excel/VBIDE `excel` backend or the pure-Go saved-workbook `file` backend.       | excel   |
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
xlflow push --backend file --json
```

## File backend

`--backend file` rebuilds the saved `.xlsm` directly from the tracked source
tree — no Excel, COM, VBIDE, or bridge process — so it works on Linux,
containers, CI runners, WSL, and remote coding agents. It preserves existing
UserForm designer storage and updates form code-behind (including `sidecar`
mode sidecars), but it cannot author a brand-new UserForm; use `--backend
excel` when a form does not yet exist in the workbook.

Because the backend replaces the saved workbook beneath any live Excel/VBE
state, it refuses to run — before changing anything — when the workbook looks
open (`~$` lock file, Windows open-state probe), when a matching xlflow
session is recorded, or when a matching WSL-side live session is detected.
Close Excel or stop the session first, or use `--backend excel`.

A successful file push means the workbook artifact was rebuilt and
structurally validated — it does not mean the VBA compiled. JSON reports
`push.backend="file"`, `push.target="saved_workbook"`,
`push.vbe_validation="not_performed"`, and warning `vbe_validation_skipped`.
Compile/run validation still requires opening the workbook in Excel (for
example through `xlflow run` or `xlflow test`).

`--changed-only` interoperates across both backends: state written by either
backend satisfies the other. `--backend file` cannot be combined with
`--session` or `--no-save`.

## Notes

> [!IMPORTANT]
> `push` runs source preflight before opening Excel so modal compile dialogs are caught as structured CLI errors whenever possible.

Projects may list an explicitly reviewed registry blocker under
`[preflight].allowed_diagnostics`. The diagnostic remains enabled and is
reported as an error by static analysis; `push` proceeds with aggregated
`preflight_diagnostic_allowed` warnings, one per waived diagnostic ID with
occurrence counts aggregated within each warning, and the later VBE compile can
still fail. Duplicate components, unreadable source, and UserForm artifact
integrity failures cannot be waived this way.

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
