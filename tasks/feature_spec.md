# UI Button Spec

## Goal

Allow agents and users to place a runnable entrypoint on a workbook sheet by creating an Excel form-control button and assigning an existing macro to its `OnAction`.

## CLI Contract

- `xlflow ui button add --sheet <name> --cell <A1> --text <caption> --macro <module.proc> [--id <id>] [--width <points>] [--height <points>] [--create-sheet] [--verify-macro]`
- `xlflow ui button list [--sheet <name>]`
- `xlflow ui button remove --id <id> [--sheet <name>]`
- `add` creates or updates an Excel form-control button named `xlflow.button.<id>`.
- `--id` is normalized to a lowercase ASCII slug. When omitted, it is derived from `--macro`.
- `--width` defaults to `160`; `--height` defaults to `40`.
- `--verify-macro` checks VBIDE for the macro before saving; without it, the button is added even if the macro cannot be inspected.

## Output Contract

JSON output adds a top-level `ui` field. `add` and `remove` return `ui.button`; `list` returns `ui.buttons`. Button objects include `id`, `name`, `sheet`, `text`, `macro`, `cell`, `left`, `top`, `width`, `height`, and `updated`.

## Failure Contract

- CLI argument errors return exit code `2`.
- Missing sheets, missing buttons, invalid cells, and missing macros under `--verify-macro` are validation failures.
- Excel COM, workbook open, PowerShell, and VBIDE access failures are environment failures.
