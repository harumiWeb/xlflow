# xlflow Diff Harness Todo

## Implementation

- [x] Add `excelize/v2` dependency.
- [x] Add workbook sheet, cell value, and formula diff backend.
- [x] Add exported VBA text diff backend for `.bas`, `.cls`, and `.frm`.
- [x] Add `xlflow diff` CLI command and argument validation.
- [x] Add top-level JSON `diff` envelope field.
- [x] Update CLI contract, README, feature spec, and ADR documentation.

## Verification

- [x] Add unit tests for CLI registration and argument validation.
- [x] Add unit tests for workbook sheet/value/formula diffs.
- [x] Add unit tests for VBA add/change and line-ending normalization.
- [x] Add output envelope JSON coverage for `diff`.
- [x] Run `go test ./...` after documentation updates.
- [x] Run `task verify`.
