# xlflow run

Run a workbook macro from the CLI.

## Usage

```bash
xlflow run [macro] [--arg <type:value>]... [--save|--save-as <path>] [--trace] [--headless|--interactive] [--session]
```

## Options and Arguments

| Option / argument      | Description                                                                    | Default             |
| ---------------------- | ------------------------------------------------------------------------------ | ------------------- |
| `macro`                | Macro entrypoint such as `Main.Run`. If omitted, config may provide the entry. | config entry        |
| `--arg <type:value>`   | Pass a typed macro argument. Repeat for multiple arguments.                    | -                   |
| `--headless`           | Run without showing Excel when possible.                                       | false               |
| `--interactive`        | Allow visible Excel interaction for dialogs or UserForms.                      | false               |
| `--session`            | Run in the managed live session workbook.                                      | false               |
| `--save`               | Save the workbook after running.                                               | false               |
| `--save-as <path>`     | Save a copy to a different workbook path.                                      | -                   |
| `--trace`              | Inject or use trace logging around the run.                                    | false               |
| `--diagnostic`         | Use diagnostic execution with stronger compile-dialog visibility.              | false               |
| `--direct`             | Run an argument-free macro without temporary harness injection.                | false               |
| `--fast`               | Use development-oriented fast run defaults.                                    | false               |
| `--gui-compile-errors` | Let Excel/VBE compile dialogs surface instead of structured diagnostics.       | false               |
| `--input <path>`       | Override workbook path for this run.                                           | configured workbook |
| `--timeout <duration>` | Maximum macro runtime before timeout.                                          | 5m0s                |
| `--keepalive`          | Write periodic progress heartbeat lines to stderr.                             | false               |

## Examples

```bash
xlflow macros --json
xlflow run Main.Run --headless --json
xlflow run Main.Run --arg string:ABC123 --session --trace --json
```

## Notes

::: tip
Discover entrypoints with `xlflow macros --json` before running macros from an agent.
:::

::: warning
Use `--interactive` only when the macro intentionally shows dialogs or UserForms. Headless automation should avoid GUI prompts.
:::

> [!IMPORTANT]
> For AI-agent debugging, prefer the default diagnostic mode and keep `--gui-compile-errors` off unless a human is watching Excel.

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "run",
  "macro": "Main.Run",
  "result": { "message": "completed" },
  "duration_ms": 1234
}
```

## Related

- [macros](./macros)
- [trace](./trace)
- [test](./test)
