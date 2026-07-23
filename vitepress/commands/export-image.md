# xlflow export-image

Export a worksheet range as a PNG image through Excel COM.

## Usage

```bash
xlflow export-image --sheet <name> --range <address> [--out <png>|--output-dir <dir>] [--session]
```

## Options and Arguments

| Option / argument    | Description                                    | Default                 |
| -------------------- | ---------------------------------------------- | ----------------------- |
| `--sheet <name>`     | Worksheet to render.                           | active/configured sheet |
| `--range <address>`  | A1 range to export.                            | used range              |
| `--out <png>`        | Exact output image path.                       | -                       |
| `--output-dir <dir>` | Directory for generated image output.          | artifacts               |
| `--name <name>`      | Base filename when using `--output-dir`.       | derived                 |
| `--format png`       | Output image format.                           | png                     |
| `--overwrite`        | Replace an existing image file.                | false                   |
| `--session`          | Render from the managed live session workbook. | false                   |

## Examples

```bash
xlflow export-image --sheet Dashboard --range A1:K30 --out artifacts/dashboard.png --json
xlflow export-image --sheet QR --range A1:AE31 --output-dir artifacts --overwrite --json
```

## Notes

::: tip
Use image exports in PRs and agent loops to review visual workbook output without opening Excel manually.
:::

::: warning
Without `--overwrite`, existing output files may cause the command to fail instead of replacing the image.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "export-image",
  "sheet": "Dashboard",
  "range": "A1:K30",
  "output": "artifacts/dashboard.png"
}
```

## Related

- [inspect](./inspect)
- [run](./run)
- [Demos](../demos/)

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow export image` when the task matches the command description above. For a goal-oriented workflow, start with the [How-to guides](../guides/) and return here for exact options.

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
