# ADR-0002: Separate Workbook and VBA Diff Backends

## Status

`accepted`

## Background

`xlflow diff` needs to show both workbook-state changes and VBA source changes. Excel workbook state is stored in OpenXML workbook parts, while editable VBA source is normally obtained through `xlflow pull` as `.bas`, `.cls`, and `.frm` files. Treating both as one backend would either overfit to Excel COM/VBIDE or hide important workbook state behind source-only diffs.

Alternative approaches considered:

- Use Excel COM/VBIDE for both workbook and VBA comparison.
- Parse `vbaProject.bin` directly from `.xlsm` files.
- Use excelize for workbook state and compare exported VBA source as normal text.

## Decision

Implement `xlflow diff` v1 with separate backends:

- Use `github.com/xuri/excelize/v2` for workbook sheet, cell value, and formula comparison.
- Compare exported VBA source directories as normalized text when `--vba-before` and `--vba-after` are provided.
- Do not parse `vbaProject.bin` or require Excel COM for v1 diff.

The command returns success when differences are found, and exposes the machine-readable result under the top-level JSON `diff` field.

## Consequences

Positive consequences:

- Workbook diffs run without Excel, COM, or VBIDE access.
- VBA diffs reuse the existing source-controlled export model.
- The v1 behavior is deterministic and easy to test in Go.

Negative consequences:

- VBA comparison requires callers to provide exported source trees.
- v1 does not compare shapes, charts, print settings, styles, comments, defined names, or recalculated semantic results.
- A future COM or OpenXML-deeper backend may need another ADR if it changes the command contract.

## Rationale

- Tests: `internal/diff` workbook and VBA diff unit tests; `internal/cli` diff command tests; `internal/output` JSON envelope tests.
- Code: `internal/diff`, `internal/cli`, `internal/output`.
- Related specs: `docs/specs/cli-contract.md`, `docs/design.md`, `tasks/feature_spec.md`.

## Supersedes

- None

## Superseded by

- None
