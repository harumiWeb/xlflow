# xlflow Diff Harness Spec

## Goal

Add `xlflow diff` so agents and humans can compare workbook state and exported VBA source after automated Excel/VBA changes.

## Behavior

- `xlflow diff <before-workbook> <after-workbook>` compares workbook sheet lists, used-range cell values, and formulas.
- Workbook inputs must use `.xlsx`, `.xlsm`, `.xltx`, or `.xltm`.
- `--vba-before <dir>` and `--vba-after <dir>` must be provided together.
- VBA comparison recursively includes `.bas`, `.cls`, and `.frm` files only.
- VBA text comparison normalizes line endings so CRLF/LF differences alone do not count as changes.
- Differences are successful command results with exit code `0`.
- Malformed arguments fail with exit code `2`.
- Unreadable workbooks or source trees fail with exit code `3`.

## Interfaces

- CLI: `xlflow [--json] diff <before-workbook> <after-workbook> [--vba-before <dir>] [--vba-after <dir>]`
- JSON: top-level `diff` field containing `summary`, `sheets`, `cells`, and `vba`.
- Plain text: existing `logs` rendering prints `no differences found` or compact human-readable diff lines.

## Verification

- Fast gate: `go test ./...` and `task verify`.
- Integration gate: create two temporary workbooks, run `xlflow diff before.xlsx after.xlsx --json`, and confirm `diff.summary.total_diffs` reflects the workbook changes.
