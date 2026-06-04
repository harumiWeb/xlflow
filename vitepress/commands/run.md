# xlflow run

Run a workbook macro from the CLI.

## Usage

```bash
xlflow run [macro] [--arg <type:value>]... [--msgbox <dialog-id=result>]... [--inputbox <dialog-id=value>]... [--filedialog <kind>:<dialog-id>=<value>]... [--ui-stream] [--save|--save-as <path>] [--trace] [--headless|--interactive] [--session]
```

## Options and Arguments

| Option / argument              | Description                                                                                | Default             |
| ------------------------------ | ------------------------------------------------------------------------------------------ | ------------------- |
| `macro`                        | Macro entrypoint such as `Main.Run`. If omitted, config may provide the entry.             | config entry        |
| `--arg <type:value>`           | Pass a typed macro argument. Repeat for multiple arguments.                                | -                   |
| `--msgbox <id=result>`         | Provide a scripted `XlflowUI.MsgBox` response. Repeat for multiple dialogs.                | -                   |
| `--inputbox <id=value>`        | Provide a scripted `XlflowUI.InputBox` response. Repeat for multiple dialogs.              | -                   |
| `--filedialog <kind:id=value>` | Provide a scripted `XlflowUI` file dialog response. Repeat for multiple values or dialogs. | -                   |
| `--ui-stream`                  | Stream resolved headless `XlflowUI` events to stderr in real time.                         | false               |
| `--headless`                   | Run without showing Excel when possible.                                                   | false               |
| `--interactive`                | Allow visible Excel interaction for dialogs or UserForms.                                  | false               |
| `--session`                    | Run in the managed live session workbook.                                                  | false               |
| `--save`                       | Save the workbook after running.                                                           | false               |
| `--save-as <path>`             | Save a copy to a different workbook path.                                                  | -                   |
| `--trace`                      | Inject or use trace logging around the run.                                                | false               |
| `--diagnostic`                 | Use diagnostic execution with stronger compile-dialog visibility.                          | false               |
| `--direct`                     | Run an argument-free macro without temporary harness injection.                            | false               |
| `--fast`                       | Use development-oriented fast run defaults.                                                | false               |
| `--gui-compile-errors`         | Let Excel/VBE compile dialogs surface instead of structured diagnostics.                   | false               |
| `--input <path>`               | Override workbook path for this run.                                                       | configured workbook |
| `--timeout <duration>`         | Maximum macro runtime before timeout.                                                      | 5m0s                |
| `--bridge <provider>`          | Select the Excel bridge provider (`auto`, `powershell`, `dotnet`).                         | auto                |

## Examples

```bash
xlflow macros --json
xlflow run Main.Run --headless --json
xlflow run Main.Run --arg string:ABC123 --session --trace --json
xlflow run Main.Run --headless --msgbox confirm-save=yes --inputbox customer-name=alice --ui-stream --json
xlflow run Main.Run --headless --filedialog get-open:source-files=C:\temp\a.txt --filedialog get-open:source-files=C:\temp\b.txt --filedialog save-as:export-path=C:\temp\out.xlsx --ui-stream --json
```

## Notes

::: tip
Discover entrypoints with `xlflow macros --json` before running macros from an agent.
:::

::: warning
Use `--interactive` only when the macro intentionally shows dialogs or UserForms. Headless automation should avoid GUI prompts.
:::

::: tip
If the macro uses `XlflowUI.MsgBox`, `XlflowUI.InputBox`, or `XlflowUI` file dialog wrappers, prefer `--msgbox`, `--inputbox`, and `--filedialog` to keep the run headless. Add `--ui-stream` when you want realtime terminal visibility into which dialog ids resolved from scripted responses versus workbook defaults.

::: tip
Supported `--filedialog` kinds are `get-open`, `file-open`, `save-as`, and `folder`. Repeat the same `kind:id=value` flag to simulate multi-select results, and use `@cancel` to simulate a cancelled dialog.
:::
:::

> [!IMPORTANT]
> For AI-agent debugging, prefer the default diagnostic mode and keep `--gui-compile-errors` off unless a human is watching Excel.

::: tip
`run` supports the .NET bridge via explicit `--bridge dotnet`. In `auto` mode, `run` uses the PowerShell bridge. With `--bridge dotnet`, macro invocation, argument passing, `--trace` temporary injection/revert, `--msgbox`/`--inputbox`/`--filedialog` response injection, `--ui-stream` module injection, and session-aware workflow all work through the .NET Excel bridge executable.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "run",
  "macro": {
    "name": "Main.Run",
    "duration_ms": 1234
  },
  "ui": {
    "events": [
      {
        "kind": "file-open",
        "dialog_id": "source-files",
        "response_source": "scripted",
        "resolved_value": "C:\\temp\\a.txt | C:\\temp\\b.txt"
      }
    ]
  }
}
```

When `--ui-stream` is enabled, xlflow also writes realtime stderr lines such as `xlflow: ui kind=file-open id=source-files source=scripted value=C:\temp\a.txt | C:\temp\b.txt`. These lines never go to stdout, so JSON stdout remains valid.

## Related

- [macros](./macros)
- [trace](./trace)
- [test](./test)
