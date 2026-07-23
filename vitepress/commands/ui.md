# xlflow ui

Manage xlflow-owned worksheet buttons.

## Usage

```bash
xlflow ui button add --sheet <sheet> --cell <cell> --text <label> --macro <entry> [--session]
xlflow ui button list [--session]
xlflow ui button remove --id <id> [--session]
```

## Options and Arguments

| Option / argument | Description                                        | Default          |
| ----------------- | -------------------------------------------------- | ---------------- |
| `button add`      | Create or update a worksheet button.               | -                |
| `button list`     | List xlflow-owned buttons.                         | -                |
| `button remove`   | Remove a managed button.                           | -                |
| `--sheet <name>`  | Target worksheet.                                  | required for add |
| `--cell <addr>`   | Anchor cell.                                       | required for add |
| `--text <label>`  | Button text.                                       | required for add |
| `--macro <entry>` | Assigned macro entrypoint.                         | required for add |
| `--verify-macro`  | Fail if the target macro is not discoverable.      | false            |
| `--session`       | Operate against the managed live session workbook. | false            |

## Examples

```bash
xlflow ui button add --sheet Home --cell B2 --text "Run" --macro Main.Run --verify-macro --json
xlflow ui button list --json
```

## Notes

::: tip
Give buttons stable ids or anchors so rerunning the command can update the intended control.
:::

When an xlflow session is active and `.xlflow/session.json` points at the configured workbook, `ui button` commands auto-reuse that matching live session workbook instead of opening a fresh hidden instance. `add` and `remove` save the live session workbook after successful mutation so the session workbook stays open; `list` is read-only and does not save.

`ui button add|list|remove` uses the `.NET` bridge on Windows in `auto` mode.

::: warning save_failed failure mode
If the workbook save fails after a successful button mutation, the command sets `"status": "error"` with error code `save_failed`. In this state the live session still holds the mutation, but the workbook file on disk may not reflect it. Run `xlflow save --session` to retry persisting the changes.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "ui",
  "ui": {
    "button": {
      "id": "36c99bbe-f155-459e-9474-215b6b7db5c4",
      "name": "xlflow.button.36c99bbe-f155-459e-9474-215b6b7db5c4",
      "sheet": "Home",
      "cell": "B2",
      "text": "Run",
      "macro": "Main.Run",
      "left": 100,
      "top": 50,
      "width": 160,
      "height": 40,
      "updated": false
    }
  }
}
```

The `updated` field indicates whether an existing button was updated (`true`) or a new button was created (`false`). This field only appears for the `button add` action.

## Related

- [macros](./macros)
- [run](./run)

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow ui` when the task matches the command description above. For a goal-oriented workflow, start with the [How-to guides](../guides/) and return here for exact options.

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
