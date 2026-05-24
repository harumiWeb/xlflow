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
