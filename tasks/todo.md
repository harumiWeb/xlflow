# xlflow Test Harness Todo

## Implementation

- [x] Add failing tests for `xlflow test` CLI registration and runner wiring.
- [x] Add failing tests for JSON `tests` passthrough and test failure exit-code classification.
- [x] Add failing tests for PowerShell test discovery, filter, and parse coverage.
- [x] Add failing tests for `XlflowAssert.bas` scaffold creation.
- [x] Implement the Go CLI, output, and Excel runner changes.
- [x] Implement `scripts/test.ps1` and shared PowerShell helper functions.
- [x] Add `XlflowAssert.bas` to project scaffolding.
- [x] Update `docs/specs/cli-contract.md` and `README.md`.

## Verification

- [x] Run focused Go tests for touched packages.
- [x] Run `go test ./...`.
- [x] Run `task verify`.
- [x] Run real Excel COM E2E for passing, failing, and filtered VBA tests when the environment allows it.
