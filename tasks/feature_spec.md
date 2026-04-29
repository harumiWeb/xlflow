# xlflow Test Harness Spec

## Goal

Add `xlflow test` so agents and CI can execute workbook VBA tests through Excel, receive stable JSON results, and use exit codes to decide whether to repair VBA or the local Excel environment.

## Behavior

- `xlflow test` opens the configured workbook and discovers tests from the workbook VBIDE state.
- Test procedures are argument-free `Sub` procedures whose names start with `Test` or end with `_Test`.
- `xlflow test --filter <name>` runs only the test whose procedure name exactly matches `<name>`.
- Duplicate discovered test names fail with `duplicate_test_name`.
- No discovered tests fail with `no_tests_found`.
- A missing filter target fails with `test_not_found`.
- Individual VBA assertion or runtime errors fail the command with exit code `1` and include per-test error details.
- Excel, COM, VBIDE, PowerShell, or script problems fail with exit code `3`.

## Interfaces

- CLI: `xlflow [--json] test [--filter <name>]`
- JSON: top-level `tests` field containing test result objects with `name`, `module`, `status`, `duration_ms`, and optional `error`.
- Scaffold: `src/modules/XlflowAssert.bas` provides `AssertEquals`.

## Verification

- Fast gate: `go test ./...` and `task verify`.
- Real Excel gate: create a disposable project, push a test module, run `xlflow test --json`, verify passing, failing, and exact filter cases.
