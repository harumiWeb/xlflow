# GUI Boundary and Human-Assisted Run Todo

## Phase 0: Spec and decision record

- [x] Define GUI boundary model in `tasks/feature_spec.md`.
- [x] Update CLI/runtime specs for run modes, inspect-gui, attach, and JSON output.
- [x] Update ADR-0001 to record GUI boundaries as an architecture policy.

## Phase 1: GUI boundary analysis

- [x] Add common `internal/gui` analyzer.
- [x] Move `lint` VB007 detection onto the common analyzer.
- [x] Add boundary metadata to lint JSON findings.
- [x] Add GUI boundary summary to `doctor`.
- [x] Update bundled xlflow skill with GUI/headless guidance.

## Phase 2: Run modes

- [x] Add `run --headless`, `run --interactive`, and `run --timeout`.
- [x] Add headless preflight before Excel starts.
- [x] Make interactive runs visible and alerts-enabled.
- [x] Return `macro_timeout` for timed-out run processes.

## Phase 3: inspect-gui

- [x] Add `xlflow inspect-gui`.
- [x] Return stable top-level `gui_boundaries` JSON.
- [x] Add human-readable boundary rendering.

## Phase 4: attach --active

- [x] Add `xlflow attach --active`.
- [x] Add `scripts/attach.ps1`.
- [x] Validate active workbook path against configured `excel.path`.

## Final verification

- [x] Run focused Go tests for changed packages.
- [x] Run PowerShell script tests.
- [x] Run full `go test ./...` with an 8-minute timeout.
- [x] Run `task verify` if full tests pass.
