# xlflow test

Discover and run workbook VBA test procedures.

## Usage

```bash
xlflow test [--filter <pattern>] [--msgbox <id=result>] [--inputbox <id=value>] [--filedialog <kind:id=value>] [--ui-stream] [--session] [--json]

```

## Options and Arguments

| Option / argument              | Description                                                           | Default |
| ------------------------------ | --------------------------------------------------------------------- | ------- |
| `--filter <pattern>`           | Run only matching test names.                                         | -       |
| `--msgbox <id=result>`         | Provide a scripted `XlflowUI.MsgBox` response. Repeat as needed.      | -       |
| `--inputbox <id=value>`        | Provide a scripted `XlflowUI.InputBox` response. Repeat as needed.    | -       |
| `--filedialog <kind:id=value>` | Provide a scripted `XlflowUI` file dialog response. Repeat as needed. | -       |
| `--ui-stream`                  | Stream resolved headless `XlflowUI` events to stderr in real time.    | false   |
| `--session`                    | Run tests in the managed live workbook.                               | false   |
| `--json`                       | Return structured test results.                                       | false   |

## Examples

```bash
xlflow test --json
xlflow test --filter Smoke --session --json
xlflow test --msgbox test-confirm=ok --inputbox test-user=alice --ui-stream --json
xlflow test --filedialog folder:export-dir=@cancel --ui-stream --json
```

## Notes

> [!IMPORTANT]
> `test` executes VBA. Use a controlled workbook state before running tests that mutate sheets or files.

::: tip
Keep VBA assertions simple and scalar so failures are easy for agents to parse.
:::

::: tip
When tests use `XlflowUI`, pass `--msgbox`, `--inputbox`, and `--filedialog` for deterministic unattended execution. Add `--ui-stream` when you want realtime confirmation of which dialog path the test took.

::: tip
Supported `--filedialog` kinds are `get-open`, `file-open`, `save-as`, and `folder`. Repeat the same `kind:id=value` flag for multi-select dialogs, and use `@cancel` to simulate a user cancellation.
:::
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "test",
  "tests": [{ "name": "TestSmoke", "status": "pass" }],
  "ui": {
    "events": [
      {
        "kind": "folder",
        "dialog_id": "export-dir",
        "response_source": "scripted",
        "resolved_value": "@cancel"
      }
    ]
  }
}
```

When `--ui-stream` is enabled, xlflow also writes realtime stderr lines such as `xlflow: ui kind=folder id=export-dir source=scripted value=@cancel` while tests are still running.

## Related

- [run](./run)
- [check](./check)
