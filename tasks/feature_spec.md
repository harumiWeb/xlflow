# Bug Fix Spec: XlflowDebug ParamArray + run compile watcher failure

## Overview

Fix two independent bugs that cause `xlflow run` to produce spurious `XlflowDebug.bas` compile errors and misclassify VBE compile control lookup failures.

## Fix A: XlflowDebug ParamArray ByRef

**Goal:** `XlflowDebug.bas` helper module must not cause additional compile errors when user code has compile errors.

**Root cause:** `defaultDebugRuntimeModule` in `internal/project/scaffold.go` defined `JoinLogMessage(ByRef Parts() As Variant)`, which cannot receive a `ParamArray` in some VBA environments.

**Fix:** Change `ByRef` to `ByVal` in `JoinLogMessage` parameter declaration. Public interface `Log(ParamArray Parts() As Variant)` is preserved.

**Regression test:** `TestXlflowDebugJoinLogMessageDoesNotForceParamArrayToByRef` in `internal/project/scaffold_test.go` — verifies the scaffolded module does not contain `ByRef Parts() As Variant`.

## Fix B: compile watcher failure misclassification

**Goal:** When `Invoke-XlflowVBECompile` cannot find the VBE compile command control, the failure must be reported as `vba_compile_failed`, not misclassified as `vbide_access_denied` or `macro_not_found`.

**Root cause:** `Invoke-XlflowVBECompile` in `internal/excel/scripts/common.ps1` catch block did not set `$result.ok = $false`, allowing `run.ps1` to route the failure incorrectly.

**Fix:** Set `$result.ok = $false` in the catch block of `Invoke-XlflowVBECompile` so callers can detect the failure.

**Regression test:** `TestInvokeXlflowVBECompileMarksFailureWhenCompileControlNotFound` in `internal/excel/scripts/scripts_test.go` — dot-sources common.ps1, mocks `Get-XlflowVBECompileControl` returning `$null`, verifies `ok=false` with error message.

## Expected Behavior After Fix

- `xlflow run` with `SampleFail` macro executes without spurious `XlflowDebug.bas` compile errors
- User code compile failure is reported as `vba_compile_failed` (structured error)
- No GUI dialog residuals block subsequent workflow
- Regression tests pass from failure to success

## E2E Verification Points

- `SampleFail` macro: `Public Sub SampleFail(): Dim x As Integer: x = "abc": End Sub`
- session-first workflow: `session start → push --fast --session --no-save → run --session → save --session → session stop`
- Confirm: no `ParamArray` compile error, no GUI dialog hang, structured failure in terminal

---

# Phase 1-2 Feature Spec: xlflow-native Test UX

## Overview

Implement Phase 1 (scaffold cleanup + XlflowAssert expansion) and Phase 2 (lifecycle hooks + inconclusive + --module/--tag filters) from `docs/design.md`.

---

## Phase 1: Scaffold Cleanup + XlflowAssert Expansion

### 1.1 Remove `tests/` Scaffold

**Goal:** Stop creating the unused `tests/` directory during `xlflow new` and `xlflow init`.

**Change:**

- File: `internal/project/scaffold.go`
- Remove `"tests"` from the `dirs` slice in `createScaffold()`.

**Validation:**

- File: `internal/project/scaffold_test.go`
- Remove `"tests"` from `TestInitScaffold` expected paths.

### 1.2 Expand `XlflowAssert.bas`

**Goal:** Provide scalar-friendly assertions aligned with the design doc.

**New public subs (all `Public Sub`, parameterless return, raise `Err` on failure):**

| Sub                  | Signature                       | Error Number          | Notes                             |
| -------------------- | ------------------------------- | --------------------- | --------------------------------- |
| `AssertEquals`       | `(expected, actual, [message])` | `vbObjectError + 513` | Existing                          |
| `AssertNotEqual`     | `(expected, actual, [message])` | `vbObjectError + 513` | New                               |
| `AssertTrue`         | `(condition, [message])`        | `vbObjectError + 513` | New                               |
| `AssertFalse`        | `(condition, [message])`        | `vbObjectError + 513` | New                               |
| `AssertFail`         | `([message])`                   | `vbObjectError + 513` | New                               |
| `AssertInconclusive` | `([message])`                   | `vbObjectError + 516` | New; used by runner to map status |
| `AssertIsNothing`    | `(value, [message])`            | `vbObjectError + 513` | New                               |
| `AssertIsNotNothing` | `(value, [message])`            | `vbObjectError + 513` | New                               |

**Private helpers:**

- `RaiseAssertFailure(message, detail)` — shared helper
- `DescribeAssertValue(value)` — existing, keep

**Validation:**

- `TestScaffoldCreatesAssertHelper` verifies all new public subs exist in the scaffolded `.bas` file.
- `TestInstallHelperModulesUsesConfiguredModuleRoot` verifies helpers are installed by `module install`.

---

## Phase 2: Lifecycle Hooks + Filters

### 2.1 Hook Discovery

**Goal:** Discover `BeforeAll`, `AfterAll`, `BeforeEach`, `AfterEach` per module.

**PowerShell Function:** `Find-XlflowModuleHooks`

```powershell
function Find-XlflowModuleHooks {
  param([string]$ModuleName, [string]$Code)
  # Returns ordered array of hook objects:
  #   @{ name = "BeforeAll";  module = $ModuleName; line = <n> }
  #   @{ name = "AfterAll";   module = $ModuleName; line = <n> }
  #   @{ name = "BeforeEach"; module = $ModuleName; line = <n> }
  #   @{ name = "AfterEach";  module = $ModuleName; line = <n> }
}
```

**Rules:**

- Match `^(?:Public\s+)?Sub\s+(BeforeAll|AfterAll|BeforeEach|AfterEach)\s*(?:\(\s*\))?\s*(?:''.*)?$`
- Same `Public` / implicit-public rule as `Find-XlflowTestProcedures`.
- Case-insensitive.

### 2.2 Runner Generation

**Goal:** Replace `Select Case` dispatch with per-test-case execution that includes hooks.

**PowerShell Function:** `New-XlflowTestRunnerCode`

**Old behavior:**

- Generates a single `RunTest(ByVal testIndex As Long)` with `Select Case`.
- Returns `Array(success, errNumber, errSource, errDescription)`.

**New behavior:**

- Generates two functions:
  1. `RunTestCase(ByVal moduleName As String, ByVal testName As String, ByVal beforeEachName As String, ByVal afterEachName As String) As Variant`
  2. `RunHook(ByVal moduleName As String, ByVal hookName As String) As Variant`
- `RunTestCase` execution order:
  1. If `beforeEachName` is not empty, call `RunHook(moduleName, beforeEachName)`.
     - If hook fails, skip test body but **still execute** `afterEachName` if present.
  2. If `beforeEach` passed, call `moduleName.testName`.
  3. If `afterEachName` is not empty, call `RunHook(moduleName, afterEachName)`.
     - `AfterEach` failure overwrites any prior success.
  4. Return `Array(success, errNumber, errSource, errDescription, statusHint)`.
     - `statusHint` is `""` for normal pass/fail, `"inconclusive"` when `errNumber = vbObjectError + 516`.

- `RunHook` is a thin wrapper that calls `moduleName.hookName` with `On Error Resume Next` and returns the same `Array(...)` shape.

### 2.3 Test Execution Loop (`test.ps1`)

**Goal:** Group tests by module, run `BeforeAll` once per module, then per-test with `BeforeEach`/`AfterEach`, then `AfterAll`.

**New flow:**

```text
foreach ($moduleGroup in $selected | Group-Object module) {
    $moduleName = $moduleGroup.Name
    $hooks = Find-XlflowModuleHooks -ModuleName $moduleName -Code (Get-XlflowCodeModuleText ...)

    # BeforeAll
    if ($hooks.BeforeAll exists) {
        $result = excel.Run("runnerName.RunHook", $moduleName, "BeforeAll")
        if failed {
            # Mark ALL tests in this module as failed with error.code = "before_all_failed"
            continue to next module
        }
    }

    foreach ($test in $moduleGroup.Group) {
        # RunTestCase(module, test, beforeEach, afterEach)
        $runResult = excel.Run("runnerName.RunTestCase", ...)
        # Map result to JSON entry
    }

    # AfterAll
    if ($hooks.AfterAll exists) {
        $result = excel.Run("runnerName.RunHook", $moduleName, "AfterAll")
        if failed {
            # Mark ALL tests in this module as failed with error.code = "after_all_failed"
        }
    }
}
```

**Failure mapping:**

| Failure              | `status`       | `error.code`         | Affected Tests           |
| -------------------- | -------------- | -------------------- | ------------------------ |
| Test body assertion  | `failed`       | `test_failed`        | Single test              |
| `BeforeAll`          | `failed`       | `before_all_failed`  | All tests in module      |
| `AfterAll`           | `failed`       | `after_all_failed`   | All tests in module      |
| `BeforeEach`         | `failed`       | `before_each_failed` | Single test              |
| `AfterEach`          | `failed`       | `after_each_failed`  | Single test (overwrites) |
| `AssertInconclusive` | `inconclusive` | `test_inconclusive`  | Single test              |

**Notes:**

- `BeforeEach` failure skips the test body but **still runs** `AfterEach` for cleanup.
- `AfterEach` failure overrides any prior status for that test.

### 2.4 `--module` / `--tag` Filter Infrastructure

**Goal:** Add CLI flags and PowerShell selection logic. Full tag metadata parsing is Phase 2 foundation for Phase 3 tag support.

**Go CLI (`internal/cli/root.go`):**

- Add `--module` string flag to `test` command.
- Add `--tag` string flag to `test` command.
- Pass both through `buildTestScriptArgs` into PowerShell parameters.

**PowerShell (`test.ps1`):**

- New params: `[string]$ModuleFilter = ""`, `[string]$TagFilter = ""`
- Pass to `Select-XlflowTests`.

**PowerShell (`common.ps1` `Select-XlflowTests`):**

- New signature: `param($Tests, [string]$Filter = "", [string]$ModuleFilter = "", [string]$TagFilter = "")`
- Apply AND logic:
  - If `$Filter` non-empty, require exact name match.
  - If `$ModuleFilter` non-empty, require exact module name match.
  - If `$TagFilter` non-empty, require any tag match (Phase 3 will implement tag extraction; Phase 2 keeps the selection contract ready).

### 2.5 `inconclusive` Terminal Output

**File:** `internal/output/output.go`

- Add `inconclusive` to the renderer summary line (`N passed, N failed, N inconclusive, N total`).
- Display inconclusive tests with `[?]` prefix.

---

## JSON Contract Updates

### Test Result Object

```json
{
  "name": "TestSmoke",
  "module": "Tests",
  "status": "passed | failed | inconclusive",
  "duration_ms": 12,
  "error": {
    "code": "test_failed | before_all_failed | after_all_failed | before_each_failed | after_each_failed | test_inconclusive",
    "message": "...",
    "source": "...",
    "number": 123
  },
  "tags": [] // Phase 3
}
```

---

## E2E Verification Plan

1. **Scaffold cleanup:** `xlflow new` in temp dir → confirm no `tests/` folder.
2. **XlflowAssert expansion:** `module install` → read `.bas` → verify all subs present.
3. **Hook happy path:** Create module with `BeforeAll`, `BeforeEach`, `TestOk`, `AfterEach`, `AfterAll` → run `xlflow test --session` → verify JSON shows all passed.
4. **BeforeEach failure:** `BeforeEach` raises error → verify test skipped, `AfterEach` still runs, status = `failed`, code = `before_each_failed`.
5. **BeforeAll failure:** `BeforeAll` raises error → verify ALL tests in module = `failed`, code = `before_all_failed`.
6. **AfterAll failure:** `AfterAll` raises error → verify ALL tests in module = `failed`, code = `after_all_failed`.
7. **Inconclusive:** `AssertInconclusive` in test → verify status = `inconclusive`, code = `test_inconclusive`.
8. **--module filter:** Create two test modules → run with `--module ModuleA` → only ModuleA tests execute.

---

## Files to Change

### Phase 1

- `internal/project/scaffold.go`
- `internal/project/scaffold_test.go`
- `docs/specs/cli-contract.md`
- `vitepress/commands/test.md`
- `vitepress/reference/project-structure.md`
- `docs/adr/NNNN-tests-located-in-src.md`

### Phase 2

- `internal/excel/scripts/common.ps1`
- `internal/excel/scripts/test.ps1`
- `internal/cli/root.go`
- `internal/excel/bridge.go` (if `buildTestScriptArgs` needs new params)
- `internal/output/output.go`
- `docs/specs/cli-contract.md`
- `vitepress/commands/test.md`
- `docs/adr/NNNN-lifecycle-hooks.md`

---

# Takt Workflow Spec: supervisor monitor support

## Goal

Align `.takt/workflows/xlflow-orchestra-high.yaml` with the added `loop_monitors` entry for the `self_review` / `fix` cycle.

## Required Changes

- Add a `fix` step that performs focused remediation after `self_review`.
- Change `self_review` so `修正が必要` transitions to `fix` instead of `implementation`.
- Add a `supervisor` persona under `.takt/facets/personas/` for loop monitor judgments.

## Behavioral Contract

- `self_review -> fix -> self_review` becomes the dedicated remediation loop.
- `fix` is narrower than `implementation`; it should address review findings with minimal scope and fresh verification.
- The `supervisor` persona judges whether repeated `self_review` / `fix` cycles show concrete progress or should abort.
- Apply the same structure to both `.takt/workflows/xlflow-orchestra-high.yaml` and `.takt/workflows/xlflow-orchestra-low.yaml`, preserving each workflow's existing model assignments.
