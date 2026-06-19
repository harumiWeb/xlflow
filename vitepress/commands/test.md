# xlflow test

Discover and run workbook VBA test procedures.

## Usage

```bash
xlflow test [--filter <pattern>] [--module <name>] [--tag <tag>] [--msgbox <id=result>] [--inputbox <id=value>] [--filedialog <kind:id=value>] [--ui-stream] [--session] [--json]

```

## Options and Arguments

| Option / argument              | Description                                                           | Default |
| ------------------------------ | --------------------------------------------------------------------- | ------- |
| `--filter <pattern>`           | Run only the test whose procedure name exactly matches the filter.    | -       |
| `--module <name>`              | Run only tests in the module whose name exactly matches the filter.   | -       |
| `--tag <tag>`                  | Run only tests tagged with the given tag.                             | -       |
| `--msgbox <id=result>`         | Provide a scripted `XlflowUI.MsgBox` response. Repeat as needed.      | -       |
| `--inputbox <id=value>`        | Provide a scripted `XlflowUI.InputBox` response. Repeat as needed.    | -       |
| `--filedialog <kind:id=value>` | Provide a scripted `XlflowUI` file dialog response. Repeat as needed. | -       |
| `--ui-stream`                  | Stream resolved headless `XlflowUI` events to stderr in real time.    | false   |
| `--session`                    | Run tests in the managed live workbook.                               | false   |
| `--bridge <provider>`          | Select the Excel bridge provider (`auto`, `powershell`, `dotnet`).    | auto    |
| `--json`                       | Return structured test results.                                       | false   |

## Test Location

Tests should live under the configured module source directory (for example `src/modules/Tests/`). This keeps tests under the same source tree as production code so `push` naturally imports them into the workbook without a separate folder convention.
Test procedure names may use Unicode VBA identifiers, so names such as `Test_集計結果が正しい` and `集計結果_Test` are discovered.

```text
src/
  modules/
    Main.bas
    Tests/
      SmokeTests.bas
      IntegrationTests.bas
```

## Tags

Add `' @Tag("name")` comment lines directly above a test sub to attach tags:

```vb
'@Tag("smoke")
Public Sub Test_CreateWorksheet()

'@Tag("integration")
'@Tag("slow")
Public Sub Test_ImportLargeFile()
```

Multiple tags are allowed. Tags are case-insensitive during filtering.

## Lifecycle Hooks

xlflow recognizes these reserved procedure names in each test module:

- `BeforeAll` — runs once before any test in the same module.
- `AfterAll` — runs once after all tests in the same module.
- `BeforeEach` — runs before each individual test in the same module.
- `AfterEach` — runs after each individual test in the same module.

All hooks must be public parameterless `Sub` procedures. If a hook fails, the affected tests are recorded as failed with a dedicated error code.

## Examples

```bash
xlflow test --json
xlflow test --filter TestSmoke --session --json
xlflow test --module SmokeTests --session --json
xlflow test --tag smoke --session --json
xlflow test --msgbox test-confirm=ok --inputbox test-user=alice --ui-stream --json
xlflow test --filedialog folder:export-dir=@cancel --ui-stream --json
```

## Notes

> [!IMPORTANT]
> `test` executes VBA. Use a controlled workbook state before running tests that mutate sheets or files.
> `test` reports progress on stderr. Interactive terminals show a spinner, while non-interactive or `--json` runs emit a single progress line so stdout stays parseable.

::: tip
Keep VBA assertions simple and scalar so failures are easy for agents to parse.
:::

::: tip
When tests use `XlflowUI`, pass `--msgbox`, `--inputbox`, and `--filedialog` for deterministic unattended execution. Add `--ui-stream` when you want realtime confirmation of which dialog path the test took.

::: tip
Supported `--filedialog` kinds are `get-open`, `file-open`, `save-as`, and `folder`. Repeat the same `kind:id=value` flag for multi-select dialogs, and use `@cancel` to simulate a user cancellation.
:::

::: tip
On Windows, `test` uses the `.NET` bridge by default in `auto` mode. `--bridge powershell` forces the legacy fallback path, and explicit `--bridge dotnet` stays strict with no implicit PowerShell fallback.
:::
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "test",
  "tests": [
    {
      "name": "TestSmoke",
      "module": "SmokeTests",
      "status": "passed",
      "duration_ms": 12,
      "tags": ["smoke"]
    },
    {
      "name": "TestDraft",
      "module": "SmokeTests",
      "status": "inconclusive",
      "duration_ms": 5,
      "tags": [],
      "error": {
        "code": "test_inconclusive",
        "message": "draft",
        "source": "XlflowAssert.AssertInconclusive",
        "number": 51332
      }
    },
    {
      "name": "TestBad",
      "module": "SmokeTests",
      "status": "failed",
      "duration_ms": 8,
      "tags": [],
      "error": {
        "code": "test_failed",
        "message": "expected <1> but got <2>",
        "source": "XlflowAssert.AssertEquals",
        "number": 51329
      }
    }
  ],
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
