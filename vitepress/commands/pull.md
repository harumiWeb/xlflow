# xlflow pull

Export workbook VBA components and form artifacts into configured source directories.

## Usage

```bash
xlflow pull [--session] [--formulas]
```

## Options and Arguments

| Option / argument | Description                                                            | Default |
| ----------------- | ---------------------------------------------------------------------- | ------- |
| `--session`       | Pull from the managed live session workbook.                           | false   |
| `--formulas`      | Also refresh worksheet formula snapshots under `formulas/` on success. | false   |
| `--json`          | Report exported files and warnings.                                    | false   |

## Examples

```bash
xlflow pull
xlflow pull --session --json
xlflow pull --formulas --json
```

## Notes

::: tip
Pull before editing if the workbook may contain newer VBA than the source tree.
:::

> [!IMPORTANT]
> UserForm Designer state and code-behind may be written to separate sidecar paths depending on project configuration.

With `[vba.line_numbers].enabled = true`, `pull` strips only xlflow-generated, fixed-width space-padded physical line labels from workbook exports, so tracked source remains unnumbered. It never treats colon labels as generated. xlflow stops safely rather than rewriting source when it encounters non-generated or mismatched numeric labels, or numeric `GoTo`, `GoSub`, or `Resume` targets. This is configuration-only behavior; `pull` has no line-number flag.

`--formulas` runs the normal VBA pull first. If it succeeds, xlflow also extracts worksheet formulas, defined names, formula references, and parse summaries into `formulas/`.

Formula extraction reads the saved workbook file directly. When using `pull --session --formulas`, save the session first if formulas were changed only in the live workbook.

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "pull",
  "workbook": "Book.xlsm",
  "written": ["src/modules/Main.bas", "src/forms/specs/UserForm1.yaml"],
  "output": {
    "formulas": {
      "dir": "formulas",
      "manifest": "formulas/manifest.json",
      "sheet_count": 2,
      "formula_region_count": 5,
      "parse_status_summary": {
        "ok": 4,
        "partial": 1,
        "failed": 0
      },
      "defined_name_count": 2
    }
  }
}
```

## Related

- [push](./push)
- [diff](./diff)
- [project structure](../reference/project-structure)

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow pull` when the task matches the command description above. For a goal-oriented workflow, start with the [How-to guides](../guides/) and return here for exact options.

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
