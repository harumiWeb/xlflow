# xlflow Trace Harness Todo

## Implementation

- [x] Add `xlflow trace inject [workbook]` CLI command.
- [x] Add `xlflow run --trace` CLI flag and script argument plumbing.
- [x] Add `XlflowTrace` VBA module generation.
- [x] Add PowerShell trace injection script.
- [x] Add run-time trace initialization, missing-module validation, log reading, and JSON/plain-text output.
- [x] Update CLI contract, README, and feature spec.

## Verification

- [x] Add unit tests for CLI registration and run trace option parsing.
- [x] Add bridge tests for trace script arguments and `trace_not_injected` exit classification.
- [x] Add PowerShell helper tests for trace module generation, run harness setup, and event parsing.
- [x] Run `go test ./...`.
- [x] Run `task verify`.
- [x] Run Excel COM E2E trace flow when Excel/VBIDE access is available.
