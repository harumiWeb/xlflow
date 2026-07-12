# xlflow test

Discover and run workbook VBA test procedures.

## Usage

```bash
xlflow test [--filter <pattern>] [--module <name>] [--tag <tag>] [--isolation none|module|test] [--no-save] [--msgbox <id=result>] [--inputbox <id=value>] [--filedialog <kind:id=value>] [--ui-stream] [--session] [--json]
xlflow test list [--module <name>] [--path <path>] --json

```

## Options and Arguments

| Option / argument              | Description                                                                                 | Default |
| ------------------------------ | ------------------------------------------------------------------------------------------- | ------- |
| `--filter <pattern>`           | Run only the test whose qualified name or unique procedure name exactly matches the filter. | -       |
| `--module <name>`              | Run only tests in the module whose name exactly matches the filter.                         | -       |
| `--tag <tag>`                  | Run only tests tagged with the given tag.                                                   | -       |
| `--isolation <mode>`           | Workbook isolation mode: `none`, `module`, or `test`.                                       | none    |
| `--no-save`                    | Do not explicitly save the workbook used for test execution.                                | false   |
| `--msgbox <id=result>`         | Provide a scripted `XlflowUI.MsgBox` response. Repeat as needed.                            | -       |
| `--inputbox <id=value>`        | Provide a scripted `XlflowUI.InputBox` response. Repeat as needed.                          | -       |
| `--filedialog <kind:id=value>` | Provide a scripted `XlflowUI` file dialog response. Repeat as needed.                       | -       |
| `--ui-stream`                  | Stream resolved headless `XlflowUI` events to stderr in real time.                          | false   |
| `--session`                    | Run tests in the managed live workbook.                                                     | false   |
| `--bridge <provider>`          | Select the Excel bridge provider (`auto`, `dotnet`).                                        | auto    |
| `--json`                       | Return structured test results.                                                             | false   |

## Workbook Isolation

By default, `xlflow test` runs against a temporary copy of the configured workbook under `.xlflow/test-runs/<run-id>/` and attempts to remove that directory after the run. Cleanup failures are reported in `test_run.cleanup` with the retained path and message when available. The configured project workbook is not opened for mutation or explicitly saved.

`--isolation none` uses one temporary workbook for the whole selected run. This is the fastest mode, but tests can share workbook state with later tests in the same run.

`--isolation module` creates a fresh workbook copy for each selected test module. Tests in one module share state, but different modules cannot affect each other. `BeforeAll` and `AfterAll` run once for that module's isolated workbook.

`--isolation test` creates a fresh workbook copy for each selected test. `BeforeAll`, `BeforeEach`, the test, `AfterEach`, and `AfterAll` all run inside that individual workbook copy.

`--session` attaches to the live managed workbook and supports only `--isolation none`. `--session --isolation module` and `--session --isolation test` fail with `unsupported_test_isolation`. `--session --no-save` prevents xlflow from explicitly saving after tests, but mutations made by VBA remain visible in the live workbook.

## Source Test Discovery

`xlflow test list --json` lists source-defined VBA tests without opening Excel or executing workbook VBA. It scans standard `.bas` modules in the configured source tree, recognizes public `Sub` procedures named `Test*` or `*_Test`, and collects `@Tag("name")`, `@ExpectedError(...)`, `@Skip`, `@Todo`, and `@TestCase(...)` comment annotations directly above each test. Class modules, UserForms, functions, private procedures, and unsupported parameterized procedures are not listed.

The JSON envelope uses `command: "test list"` and returns discovery data under `tests`:

```json
{
  "status": "ok",
  "command": "test list",
  "tests": {
    "root": "src",
    "summary": {
      "files": 1,
      "tests": 1
    },
    "items": [
      {
        "id": "SmokeTests.TestSmoke",
        "module": "SmokeTests",
        "name": "TestSmoke",
        "qualified_name": "SmokeTests.TestSmoke",
        "qualified_procedure": "SmokeTests.TestSmoke",
        "source_path": "src/modules/SmokeTests.bas",
        "line": 5,
        "procedure_line": 5,
        "tags": ["smoke"],
        "status_hint": "skipped",
        "skip": {
          "reason": "Requires Access"
        },
        "expected_error": {
          "number": 5,
          "description": "Invalid value",
          "source": "ParserModule"
        }
      }
    ]
  }
}
```

This command is intended for editor integrations such as VS Code Testing API discovery. Running tests remains the responsibility of `xlflow test`.

## Test Identity and Filtering

Parameterless tests have a stable qualified identifier in `<Module>.<Procedure>` form. Parameterized tests have one identifier per case in `<Module>.<Procedure>[<case>]` form. JSON output includes the executable identifier as both `id` and `qualified_name`, and includes the owning procedure as `qualified_procedure`.

```bash
xlflow test --filter SmokeTests.TestSmoke --session --json
```

For parameterized tests, filter by the qualified procedure to run all cases, or by the exact case ID to run one case:

```bash
xlflow test --filter MathTests.Test_Add --json
xlflow test --filter "MathTests.Test_Add[1,2,3]" --json
```

Quote exact case filters in shells that treat brackets specially. Unqualified procedure names still work when they identify exactly one discovered item after any `--module` and `--tag` filters. If the same procedure name exists in multiple modules or expands to multiple cases, use the qualified procedure or exact case ID. Ambiguous unqualified filters fail with `ambiguous_test_name` and list matching qualified names in `error.details.matches`.

## Parameterized Tests

Use `@TestCase` directly above a public test procedure with `ByVal` scalar parameters:

```vb
'@TestCase(1, 2, 3)
'@TestCase("cancels-out"; -1, 1, 0)
Public Sub Test_Add( _
    ByVal leftValue As Long, _
    ByVal rightValue As Long, _
    ByVal expected As Long)

    XlflowAssert.AssertEquals expected, Add(leftValue, rightValue)
End Sub
```

Each annotation becomes an independent test case. Supported argument literals are integer, floating-point/scientific notation, string literals with doubled VBA quotes, `True`, `False`, `Empty`, `Null`, and VBA date literals such as `#2026-07-12#`. Supported parameter types are `Boolean`, `Byte`, `Integer`, `Long`, `LongLong`, `LongPtr`, `Single`, `Double`, `Currency`, `Date`, `String`, and `Variant`.

xlflow does not evaluate constants, enum members, function calls, member expressions, arrays, object creation, workbook references, or object parameters. `ByRef`, `Optional`, and `ParamArray` parameters are rejected during discovery with `invalid_test_case`.

Unnamed case IDs use a lexical canonical form. Whitespace and string escaping are normalized, but semantically equivalent numeric spellings such as `1.0` and `1#` are not currently treated as identical.

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

## Expected VBA Errors

Add `' @ExpectedError(...)` directly above a test when the test should pass only if the test body raises a specific VBA error:

```vb
'@ExpectedError(5)
Public Sub Test_InvalidArgument()
    ParseValue ""
End Sub

'@ExpectedError(5, "Invalid value", "ParserModule")
Public Sub Test_InvalidArgument_Detail()
    ParseValue ""
End Sub
```

Supported forms are `@ExpectedError(number)`, `@ExpectedError(number, description)`, and `@ExpectedError(number, description, source)`. The error number must match exactly. Description matching is exact and case-sensitive. Source matching is exact and case-insensitive. Substring and regular-expression matching are not supported.

Only errors raised by the test procedure body can satisfy `@ExpectedError`. `BeforeAll`, `BeforeEach`, `AfterEach`, and `AfterAll` failures remain hook failures, and `AfterEach` / `AfterAll` failures still fail the test even after the expected error was raised. `AssertInconclusive` remains `inconclusive` and never satisfies `@ExpectedError`.

Malformed metadata, multiple `@ExpectedError` annotations on one test, unsupported argument counts, non-numeric error numbers, malformed string literals, and `@ExpectedError` on non-test procedures fail discovery with `invalid_test_metadata`.

## Skipped and Todo Tests

Use `@Skip` for tests that should remain visible but cannot execute in the current environment, and `@Todo` for planned behavior that is not ready to execute:

```vb
'@Skip("Requires Microsoft Access")
Public Sub Test_AccessImport()
End Sub

'@Todo("Exporter implementation is pending")
Public Sub Test_NewExporter()
End Sub
```

The reason is recommended but optional, so bare `@Skip` and `@Todo` are accepted. `@Skip()` and `@Todo()` are invalid; use a quoted string or omit the parentheses.

Skipped and todo tests are discovered, selectable by qualified name/module/tag, and returned in execution results with `status: "skipped"` or `status: "todo"` plus optional `reason`. They do not invoke the test body, `BeforeEach`, or `AfterEach`, and they do not fail the command by default. `BeforeAll` and `AfterAll` run only when the selected module has at least one executable test. Duplicate `@Skip`, duplicate `@Todo`, or combining both annotations on one test fails with `invalid_test_metadata`.

## Lifecycle Hooks

xlflow recognizes these reserved procedure names in each test module:

- `BeforeAll` — runs once before any test in the same module.
- `AfterAll` — runs once after all tests in the same module.
- `BeforeEach` — runs before each individual test in the same module.
- `AfterEach` — runs after each individual test in the same module.

All hooks must be public parameterless `Sub` procedures. If a hook fails, the affected tests are recorded as failed with a dedicated error code.

## COM Object Cleanup

When test code opens another workbook with `GetObject(path)` or `Application.Workbooks.Open(path)`, close it and release the object reference immediately. Otherwise Excel can keep the file locked after `wb.Close False`, and later hooks may fail with VBA error 70 (`Permission denied`) while deleting or overwriting the file.

```vb
Public Sub Test_ReadsOutputWorkbook()
    Dim wb As Object
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo Cleanup
    Set wb = GetObject(outputPath)

    ' assertions...

Cleanup:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description

    If Not wb Is Nothing Then
        On Error Resume Next
        wb.Close False
        Set wb = Nothing
        On Error GoTo 0
    End If

    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
End Sub
```

Best-effort hook cleanup with `On Error Resume Next` can reduce cascading failures, but it does not replace releasing workbook references in the test that opened them.

## Examples

```bash
xlflow test list --json
xlflow test --json
xlflow test --isolation module --json
xlflow test --isolation test --filter SmokeTests.TestSmoke --json
xlflow test --filter SmokeTests.TestSmoke --session --no-save --json
xlflow test --msgbox test-confirm=ok --inputbox test-user=alice --ui-stream --json
xlflow test --filedialog folder:export-dir=@cancel --ui-stream --json
```

Typical terminal output distinguishes executed and non-executed statuses:

```text
PASS SmokeTests.TestSmoke
SKIP AccessTests.Test_AccessImport: Requires Microsoft Access
TODO ExportTests.Test_NewExporter: Exporter implementation is pending
? DraftTests.TestDraft: inconclusive
FAIL SmokeTests.TestBad: expected <110> but got <100>
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
On Windows, `test` uses the `.NET` bridge in `auto` mode.
:::
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "test",
  "test_run": {
    "isolation": "none",
    "session": false,
    "temporary_workbook": true,
    "source_workbook": "build/Book.xlsm",
    "workbook_saved": false,
    "cleanup": {
      "status": "completed"
    }
  },
  "tests": [
    {
      "id": "SmokeTests.TestSmoke",
      "qualified_name": "SmokeTests.TestSmoke",
      "name": "TestSmoke",
      "module": "SmokeTests",
      "status": "passed",
      "duration_ms": 12,
      "tags": ["smoke"]
    },
    {
      "id": "SmokeTests.TestInvalidArgument",
      "qualified_name": "SmokeTests.TestInvalidArgument",
      "name": "TestInvalidArgument",
      "module": "SmokeTests",
      "status": "passed",
      "duration_ms": 7,
      "tags": [],
      "expected_error": {
        "number": 5
      },
      "observed_error": {
        "number": 5,
        "source": "ParserModule",
        "message": "Invalid value"
      }
    },
    {
      "id": "SmokeTests.TestDraft",
      "qualified_name": "SmokeTests.TestDraft",
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
      "id": "AccessTests.Test_AccessImport",
      "qualified_name": "AccessTests.Test_AccessImport",
      "name": "Test_AccessImport",
      "module": "AccessTests",
      "status": "skipped",
      "duration_ms": 0,
      "tags": [],
      "reason": "Requires Microsoft Access"
    },
    {
      "id": "ExportTests.Test_NewExporter",
      "qualified_name": "ExportTests.Test_NewExporter",
      "name": "Test_NewExporter",
      "module": "ExportTests",
      "status": "todo",
      "duration_ms": 0,
      "tags": [],
      "reason": "Exporter implementation is pending"
    },
    {
      "id": "SmokeTests.TestBad",
      "qualified_name": "SmokeTests.TestBad",
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
