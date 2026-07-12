# xlflow Testing Reference

This document describes how to write, organize, run, and debug VBA tests with xlflow. Load it when the task involves creating, fixing, or reasoning about workbook-side test code.

## Test Discovery

xlflow discovers tests by scanning every standard module in the workbook for `Public Sub` (or plain `Sub`) procedures whose names match:

- `Test*`
- `*_Test`

Parameterless tests run once. Tests with `ByVal` scalar parameters must declare one or more `@TestCase` annotations directly above the procedure; each annotation becomes a separate executable case. Unsupported parameterized tests fail discovery with `invalid_test_case`. `Private Sub` and `Function` procedures are ignored during discovery.

Tests can live in any standard module. The recommended layout is `src/modules/Tests/` with one module per test suite (e.g., `TestOrders.bas`). There is no separate `tests/` directory in the scaffold; tests are source modules like everything else.

## Parameterized Tests

Use `@TestCase` for data-driven tests:

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

Unnamed cases use argument literals in the ID, for example `MathTests.Test_Add[1,2,3]`. Named cases use the string before the semicolon, for example `MathTests.Test_Add[cancels-out]`; the name is not passed to the VBA procedure.

Supported literals are integer, floating-point/scientific notation, string, `True`, `False`, `Empty`, `Null`, and VBA date literals such as `#2026-07-12#`. Supported parameter types are `Boolean`, `Byte`, `Integer`, `Long`, `LongLong`, `LongPtr`, `Single`, `Double`, `Currency`, `Date`, `String`, and `Variant`.

Do not use constants, enum members, function calls, member expressions, arrays, object creation, workbook references, object parameters, `ByRef`, `Optional`, or `ParamArray` in parameterized tests. xlflow rejects these during discovery rather than evaluating arbitrary VBA expressions.

## Naming Conventions

Use descriptive, action-oriented names that read like assertions:

```vb
Public Sub Test_TotalPrice_IncludesTax()
Public Sub Test_ReportHeader_HasCorrectFormat()
```

Avoid vague names such as `Test1`, `TestA`, or `TestStuff`.

## XlflowAssert API

Tests should use `XlflowAssert` helpers instead of raw `Err.Raise` so failures are surfaced consistently in JSON and terminal output.

```vb
XlflowAssert.AssertEquals expected, actual, "optional context message"
XlflowAssert.AssertNotEqual forbidden, actual, "optional context message"
XlflowAssert.AssertTrue condition, "optional context message"
XlflowAssert.AssertFalse condition, "optional context message"
XlflowAssert.AssertFail "unconditional failure message"
XlflowAssert.AssertInconclusive "reason this test is not ready"
XlflowAssert.AssertIsNothing objectRef, "optional context message"
XlflowAssert.AssertIsNotNothing objectRef, "optional context message"
```

Important constraints:

- `AssertEquals` / `AssertNotEqual` support **scalar values only**. Do not pass objects or arrays; compare scalar properties such as `Range.Value2`.
- `AssertIsNothing` / `AssertIsNotNothing` require an object reference. Pass `Nothing` explicitly when testing for it.
- `AssertInconclusive` raises a dedicated error number (`vbObjectError + 516`). xlflow maps this to the `inconclusive` status instead of `failed`.

## Tags

Add `' @Tag("name")` comment lines directly above a test sub to attach tags:

```vb
'@Tag("smoke")
Public Sub Test_CreateWorksheet()

'@Tag("integration")
'@Tag("slow")
Public Sub Test_ImportLargeFile()
```

Multiple tags are allowed. Tags are case-insensitive during filtering. There is no predefined tag list; use whatever names match your project conventions.

## Expected VBA Errors

Use `' @ExpectedError(...)` directly above a test when the test should pass only if the test body raises a specific VBA error:

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

Supported forms are `@ExpectedError(number)`, `@ExpectedError(number, description)`, and `@ExpectedError(number, description, source)`. The error number matches exactly, description matches exactly and case-sensitively, and source matches exactly but case-insensitively.

Only the test body can satisfy `@ExpectedError`. Hook failures (`BeforeAll`, `BeforeEach`, `AfterEach`, `AfterAll`) keep their hook-specific failure codes, and cleanup failures still fail the test. `AssertInconclusive` remains `inconclusive` even if its internal error number matches `@ExpectedError`.

Prefer `@ExpectedError` for ordinary error-path tests because VBA has no ergonomic `AssertThrows` callback pattern. Use a manual `On Error GoTo` test only when the test must recover from the error and then assert additional workbook or object state.

Malformed `@ExpectedError` metadata, multiple expected-error annotations on one procedure, non-numeric error numbers, unsupported argument counts, malformed string literals, or attaching the annotation to a non-test procedure fail discovery with `invalid_test_metadata`. Fix metadata before changing production VBA.

## Skipped and Todo Tests

Use `@Skip` for tests that are intentionally not executable in the current environment, and `@Todo` for planned behavior that should stay visible before it is ready to run:

```vb
'@Skip("Requires Microsoft Access")
Public Sub Test_AccessImport()
End Sub

'@Todo("Exporter implementation is pending")
Public Sub Test_NewExporter()
End Sub
```

Bare `@Skip` and `@Todo` are accepted, but a reason is preferred. `@Skip()` and `@Todo()` are invalid. Skipped and todo tests are discoverable and selectable by qualified name, module, and tag, but they do not invoke the test body, `BeforeEach`, or `AfterEach`. `BeforeAll` and `AfterAll` run only when the selected module has at least one executable test. Duplicate `@Skip`, duplicate `@Todo`, or combining both annotations on one test fails with `invalid_test_metadata`.

## Lifecycle Hooks

Each standard module can optionally define up to four hook subs. They must take **no arguments** and match these exact names:

```vb
Public Sub BeforeAll()
Public Sub AfterAll()
Public Sub BeforeEach()
Public Sub AfterEach()
```

Hook behavior:

| Hook         | Runs when                                | Failure impact                                                                                                                                               |
| ------------ | ---------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `BeforeAll`  | Once before the first test in the module | All tests in the module are marked `failed` with `before_all_failed`                                                                                         |
| `AfterAll`   | Once after the last test in the module   | All tests that were `passed` or `inconclusive` in that module are overwritten to `failed` with `after_all_failed`                                            |
| `BeforeEach` | Before every test in the module          | The individual test is marked `failed` with `before_each_failed`; the test body is skipped                                                                   |
| `AfterEach`  | After every test in the module           | The individual test is marked `failed` with `after_each_failed`; if the test body already failed, the `after_each` failure is still reported in `error.code` |

`BeforeEach` failure skips the test body, but `AfterEach` still runs for cleanup. `AfterAll` runs even when some executable tests in the module failed. Skipped and todo tests are not hook targets.

### Hook state sharing

Because VBA standard-module-level variables persist across `Application.Run` invocations within the same Excel process, you can use module-level variables to share state between hooks and tests:

```vb
Attribute VB_Name = "TestOrders"
Option Explicit

Private mSetupDone As Boolean

Public Sub BeforeAll()
    mSetupDone = True
End Sub

Public Sub BeforeEach()
    ' Reset per-test state
End Sub

Public Sub AfterEach()
    ' Cleanup per-test state
End Sub

Public Sub Test_Setup_Ran()
    XlflowAssert.AssertTrue mSetupDone, "BeforeAll should have run"
End Sub
```

## COM Object Cleanup in Tests

When a test opens an external workbook, always release the COM reference immediately after closing it. `GetObject(path)` and `Application.Workbooks.Open(path)` return workbook COM proxies that can keep the file locked even after `wb.Close False` if the VBA variable still holds the object reference.

Required pattern:

```vb
Set wb = GetObject(path)
wb.Close False
Set wb = Nothing

Set wb = Application.Workbooks.Open(path)
wb.Close False
Set wb = Nothing
```

Use a cleanup block when assertions or runtime errors can happen before the close:

```vb
Public Sub Test_ReadsExternalWorkbook()
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

Missing `Set wb = Nothing` often shows up away from the leaking test. Tests may pass individually but fail as a suite when `BeforeAll`, `AfterAll`, `BeforeEach`, or `AfterEach` later tries to delete or overwrite the same file. VBA error 70, `Permission denied`, or a localized write-denied message during cleanup is a strong signal to check for an unreleased workbook reference in an earlier test.

## Running Tests

### Basic usage

```bash
xlflow test --json
xlflow test --session --json
xlflow test --isolation module --json
xlflow test --isolation test --filter TestOrders.Test_CreateWorksheet --json
```

Plain non-session `xlflow test` runs against temporary workbook copies under `.xlflow/test-runs/<run-id>/` and attempts best-effort cleanup after execution. Cleanup failures are surfaced through `test_run.cleanup.status`, with `path` and `message` when applicable. Prefer `--session` during normal AI-agent development when the live managed workbook is the intended target. Session mode supports `--isolation none`; `module` and `test` isolation are non-session modes.

### Filtering

```bash
# Exact qualified test name
xlflow test --filter TestOrders.Test_CreateWorksheet --session --json

# Module filter (all tests in one module)
xlflow test --module TestOrders --session --json

# Tag filter
xlflow test --tag smoke --session --json
```

Filters can be combined. `--filter` is exact match on a qualified test name such as `TestOrders.Test_CreateWorksheet`; an unqualified procedure name works only when it resolves to exactly one test. `--module` is exact match on module name. `--tag` matches if the test has at least one matching tag (case-insensitive).

### Inconclusive tests

Tests that call `XlflowAssert.AssertInconclusive` are reported with `[?]` in terminal output and `"status": "inconclusive"` in JSON. They do not count as failures, but they also do not count as passes.

Use `inconclusive` only when the test body actually ran and could not determine a result. Use `@Skip` or `@Todo` for tests that should not execute.

### Session workflow

For repeated test-and-fix cycles, use the session-first pattern:

```bash
xlflow session start --json
xlflow push --fast --session --no-save --json
xlflow test --session --json
# ... edit source ...
xlflow push --fast --session --no-save --json
xlflow test --session --json
xlflow save --session --json
xlflow session stop --json
```

## Test Output and Failure Diagnosis

### Terminal output

```text
PASS Test_CreateWorksheet
SKIP Test_AccessImport: Requires Microsoft Access
TODO Test_NewExporter: Exporter implementation is pending
FAIL Test_TotalPrice: expected <110> but got <100>
? Test_ImportLargeFile: inconclusive
```

### JSON output

Each test entry includes:

- `name`, `module`, `status` (`passed` | `failed` | `skipped` | `todo` | `inconclusive`)
- `duration_ms`
- `reason` for `skipped` and `todo` results when provided
- `expected_error` when `@ExpectedError` is present
- `observed_error` when an expected-error test body raised an error
- `error.code` (`test_failed`, `test_inconclusive`, `expected_error_mismatch`, `before_all_failed`, `after_all_failed`, `before_each_failed`, `after_each_failed`)
- `error.message`, `error.source`, `error.number`

When `status` is `failed`, inspect `error.code` and `error.message` first. The message comes from `XlflowAssert` helpers or the VBA runtime error description.

For `expected_error_mismatch`, inspect `expected_error`, `observed_error`, and `error.message` together:

- If `observed_error` is absent, the test body did not raise an error.
- If `observed_error.number` differs, the wrong VBA error path ran.
- If only description or source differs, decide whether the implementation should raise the documented detail or the test metadata is too strict.
- If the failure code is a hook code, do not treat it as an expected-error mismatch; fix setup/cleanup first.

### Common failure codes

| `error.code`              | Meaning                                                  | Action                                                                       |
| ------------------------- | -------------------------------------------------------- | ---------------------------------------------------------------------------- |
| `test_failed`             | Assertion failed or runtime error in test body           | Read `error.message` and `error.source`; fix the test or the code under test |
| `test_inconclusive`       | `AssertInconclusive` was called                          | Use when an executed test cannot determine a result                          |
| `expected_error_mismatch` | `@ExpectedError` was missing, different, or too specific | Compare `expected_error`, `observed_error`, and `error.message`              |
| `before_all_failed`       | `BeforeAll` raised an error                              | Fix the `BeforeAll` setup logic; no module test body ran                     |
| `after_all_failed`        | `AfterAll` raised an error                               | Fix the `AfterAll` cleanup logic; the test body may have passed              |
| `before_each_failed`      | `BeforeEach` raised an error                             | Fix `BeforeEach`; the test body was skipped                                  |
| `after_each_failed`       | `AfterEach` raised an error                              | Fix `AfterEach`; the test body may have passed                               |

### Debugging failing tests

1. Reproduce with the exact failing filter:
   ```bash
   xlflow test --filter <Module.TestName> --session --json
   ```
2. If the failure is a runtime error rather than an assertion, add `XlflowDebug.Log` calls around the suspected path and rerun:
   ```bash
   xlflow test --filter <Module.TestName> --session --json
   ```
   Read the `debug` array in JSON output.
3. If the error originates in workbook code (not the test itself), add targeted `XlflowDebug.Log` calls if needed and run the relevant macro as normal:
   ```bash
   xlflow run <MacroName> --session --json
   ```
   Use [debugging.md](debugging.md) when the failing location is still unclear.
4. If the test uses `XlflowUI` dialogs, ensure the matching `--msgbox`, `--inputbox`, or `--filedialog` responses are passed to `xlflow test`.

## Best Practices for AI Agents

- **Write tests before or alongside behavior changes**, not after. Tests are the primary correctness signal for VBA work.
- **Keep tests independent**. A test should not depend on the order of other tests. Use `BeforeEach` / `AfterEach` for isolation.
- **Use `BeforeAll` sparingly**. It is appropriate for expensive one-time setup (e.g., loading a large fixture). Prefer `BeforeEach` for state reset.
- **Do not use `Application` state assumptions in tests**. Avoid `ActiveSheet`, `Selection`, and `ActiveWorkbook`. Create explicit worksheet references.
- **Tag slow or fragile tests** with `' @tag slow` or `' @tag flaky` so CI or quick-check runs can exclude them.
- **Use `@Todo`** for planned tests that should stay visible but not execute yet. Use `AssertInconclusive` only when the test body must run before it can determine a result.
- **Use `@ExpectedError` for error paths** instead of hand-written error handlers when the only assertion is that a specific VBA error should be raised.
- **Remove E2E-only test modules** before committing. If you create intentional-failure modules to verify hook behavior, do not check them into version control.
- **Use `xlflow generate test <ModuleName>`** to scaffold a new test module with hook stubs and a sample test sub. This is faster than writing the boilerplate by hand and keeps the module structure consistent.
- **Run `xlflow lint --json` after editing test source**. xlflow lint validates both production and test VBA.
