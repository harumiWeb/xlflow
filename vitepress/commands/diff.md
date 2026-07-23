# xlflow diff

Compare workbook files and optionally exported VBA source trees.

## Usage

```bash
xlflow diff <before.xlsm> <after.xlsm> [--vba-before <dir>] [--vba-after <dir>]
```

## Options and Arguments

| Option / argument    | Description                               | Default  |
| -------------------- | ----------------------------------------- | -------- |
| `before.xlsm`        | Baseline workbook.                        | required |
| `after.xlsm`         | Workbook to compare.                      | required |
| `--vba-before <dir>` | Baseline exported VBA source directory.   | -        |
| `--vba-after <dir>`  | Comparison exported VBA source directory. | -        |
| `--json`             | Return structured diff counts and paths.  | false    |

## Examples

```bash
xlflow diff before.xlsm after.xlsm --json
xlflow diff before.xlsm after.xlsm --vba-before before/src --vba-after after/src --json
```

Workbook cell diff supports OOXML workbook files (`.xlsx`, `.xlsm`, `.xltx`, `.xltm`). `.xlsb` workbook cell diff is not supported and fails with `workbook_format_unsupported`; compare exported VBA source separately when reviewing `.xlsb` projects.

## Notes

> [!IMPORTANT]
> A successful comparison can still report differences. Inspect the JSON summary instead of treating exit code `0` as no changes.

::: tip
Use `pull` before `diff` when you need workbook and source changes in one review.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "diff",
  "summary": { "workbook_diffs": 2, "vba_diffs": 1, "total_diffs": 3 }
}
```

## Related

- [pull](./pull)
- [JSON output](../reference/json-output)

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow diff` when the task matches the command description above. For a goal-oriented workflow, start with the [How-to guides](../guides/) and return here for exact options.

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
