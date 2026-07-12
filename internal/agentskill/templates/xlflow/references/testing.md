# xlflow Testing Reference

This document describes how to write, organize, run, and debug VBA tests with xlflow. Load it when the task involves creating, fixing, or reasoning about workbook-side test code.

## Test Discovery

xlflow discovers tests by scanning every standard module in the workbook for `Public Sub` (or plain `Sub`) procedures whose names match:

- `Test*`
- `*_Test`

Test procedures must take **no arguments**. Subs with arguments, `Private Sub`, and `Function` procedures are ignored during discovery.

Tests can live in any standard module. The recommended layout is `src/modules/Tests/` with one module per test suite (e.g., `TestOrders.bas`). There is no separate `tests/` directory in the scaffold; tests are source modules like everything else.

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

`BeforeEach` failure skips the test body, but `AfterEach` still runs for cleanup. `AfterAll` runs even when some tests in the module failed.

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
```

Prefer `--session` during normal AI-agent development to reuse an open workbook.

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
FAIL Test_TotalPrice: expected <110> but got <100>
? Test_ImportLargeFile: inconclusive
```

### JSON output

Each test entry includes:

- `name`, `module`, `status` (`passed` | `failed` | `inconclusive`)
- `duration_ms`
- `error.code` (`test_failed`, `test_inconclusive`, `before_all_failed`, `after_all_failed`, `before_each_failed`, `after_each_failed`)
- `error.message`, `error.source`, `error.number`

When `status` is `failed`, inspect `error.code` and `error.message` first. The message comes from `XlflowAssert` helpers or the VBA runtime error description.

### Common failure codes

| `error.code`         | Meaning                                        | Action                                                                       |
| -------------------- | ---------------------------------------------- | ---------------------------------------------------------------------------- |
| `test_failed`        | Assertion failed or runtime error in test body | Read `error.message` and `error.source`; fix the test or the code under test |
| `test_inconclusive`  | `AssertInconclusive` was called                | Expected if the test is intentionally incomplete                             |
| `before_all_failed`  | `BeforeAll` raised an error                    | Fix the `BeforeAll` setup logic; no module test body ran                     |
| `after_all_failed`   | `AfterAll` raised an error                     | Fix the `AfterAll` cleanup logic; the test body may have passed              |
| `before_each_failed` | `BeforeEach` raised an error                   | Fix `BeforeEach`; the test body was skipped                                  |
| `after_each_failed`  | `AfterEach` raised an error                    | Fix `AfterEach`; the test body may have passed                               |

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
- **Use `AssertInconclusive`** for tests that describe desired behavior not yet implemented, rather than commenting them out.
- **Remove E2E-only test modules** before committing. If you create intentional-failure modules to verify hook behavior, do not check them into version control.
- **Use `xlflow generate test <ModuleName>`** to scaffold a new test module with hook stubs and a sample test sub. This is faster than writing the boilerplate by hand and keeps the module structure consistent.
- **Run `xlflow lint --json` after editing test source**. xlflow lint validates both production and test VBA.
