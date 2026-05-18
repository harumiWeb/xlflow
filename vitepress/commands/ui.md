# xlflow ui

Manage xlflow-owned worksheet buttons.

## Usage

```bash
xlflow ui button add --sheet <sheet> --cell <cell> --text <label> --macro <entry>
xlflow ui button list
xlflow ui button remove --id <id>
```

## Options and Arguments

| Option / argument | Description                                   | Default          |
| ----------------- | --------------------------------------------- | ---------------- |
| `button add`      | Create or update a worksheet button.          | -                |
| `button list`     | List xlflow-owned buttons.                    | -                |
| `button remove`   | Remove a managed button.                      | -                |
| `--sheet <name>`  | Target worksheet.                             | required for add |
| `--cell <addr>`   | Anchor cell.                                  | required for add |
| `--text <label>`  | Button text.                                  | required for add |
| `--macro <entry>` | Assigned macro entrypoint.                    | required for add |
| `--verify-macro`  | Fail if the target macro is not discoverable. | false            |

## Examples

```bash
xlflow ui button add --sheet Home --cell B2 --text "Run" --macro Main.Run --verify-macro --json
xlflow ui button list --json
```

## Notes

::: tip
Give buttons stable ids or anchors so rerunning the command can update the intended control.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "ui button add",
  "button": { "sheet": "Home", "text": "Run", "macro": "Main.Run" }
}
```

## Related

- [macros](./macros)
- [run](./run)
