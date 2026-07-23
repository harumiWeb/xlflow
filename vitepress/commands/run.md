# xlflow run

Run a workbook macro from the CLI.

## Usage

```bash
xlflow run [macro] [--push] [--arg <type:value>]... [--msgbox <dialog-id=result>]... [--inputbox <dialog-id=value>]... [--filedialog <kind>:<dialog-id>=<value>]... [--ui-stream] [--save|--save-as <path>] [--headless|--interactive] [--session]
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
| `--push`                       | Import source VBA into the configured workbook before running the macro.                   | false               |
| `--headless`                   | Run without showing Excel when possible.                                                   | false               |
| `--interactive`                | Allow visible Excel interaction for dialogs or UserForms.                                  | false               |
| `--session`                    | Run in the managed live session workbook.                                                  | false               |
| `--save`                       | Save the workbook after running.                                                           | false               |
| `--save-as <path>`             | Save a copy to a different workbook path.                                                  | -                   |
| `--diagnostic`                 | Use diagnostic execution with stronger compile-dialog visibility.                          | false               |
| `--direct`                     | Run an argument-free macro without temporary harness injection.                            | false               |
| `--fast`                       | Use development-oriented fast run defaults.                                                | false               |
| `--gui-compile-errors`         | Let Excel/VBE compile dialogs surface instead of structured diagnostics.                   | false               |
| `--input <path>`               | Override workbook path for this run.                                                       | configured workbook |
| `--timeout <duration>`         | Maximum macro runtime before timeout.                                                      | 5m0s                |
| `--bridge <provider>`          | Select the Excel bridge provider (`auto`, `dotnet`).                                       | auto                |

## Examples

```bash
xlflow macros --json
xlflow run Main.Run --push --json
xlflow run Main.Run --headless --json
xlflow run Main.Run --arg string:ABC123 --session --json
xlflow run Main.Run --headless --msgbox confirm-save=yes --inputbox customer-name=alice --ui-stream --json
xlflow run Main.Run --headless --filedialog get-open:source-files=C:\temp\a.txt --filedialog get-open:source-files=C:\temp\b.txt --filedialog save-as:export-path=C:\temp\out.xlsx --ui-stream --json
```

## Notes

::: tip
Discover entrypoints with `xlflow macros --json` before running macros from an agent.
:::

::: tip
Use `xlflow run --push ...` when you want the command to import edited source files before executing the macro. With `--session`, xlflow pushes into the live session workbook without saving to disk first, then runs against that same session.
:::

::: warning
`--push` targets the configured project workbook, so it cannot be combined with `--input`.
:::

::: tip
When a macro fails and your source files are newer than the workbook, `run` reports that the macro may not have been pushed yet and suggests `xlflow push` or `xlflow run --push`.
:::

::: warning
Use `--interactive` only when the macro intentionally shows dialogs or UserForms. Headless automation should avoid GUI prompts.
:::

::: warning
If `run` times out after Excel work begins, VBA may still be running. xlflow
publishes workbook recovery state before releasing the normal lock. Follow-up
workbook commands fail with `workbook_recovery_required`; `--wait` does not
resolve it. Inspect `recovery` in JSON and use `session stop --discard`,
`process cleanup`, or `recovery clear` as directed.
:::

::: tip
For VBA-internal debugging, add `XlflowDebug.Log` in workbook code and inspect `debug.events` in `xlflow run --json`.
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
On Windows, `run` uses the `.NET` bridge in `auto` mode.
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
- [test](./test)
- [recovery](./recovery)

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow run` when the task matches the command description above. For a goal-oriented workflow, start with the [How-to guides](../guides/) and return here for exact options.

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
