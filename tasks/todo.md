# xlflow MVP Quality Hardening Todo

## Completed (Tasks 1-3)

- [x] Update `docs/specs/cli-contract.md` and `README.md` to reflect verified behavior.
- [x] Add regression tests for document-module normalization and malformed-header handling; harden `scripts/common.ps1`.
- [x] Add `task verify` and document local verify vs `tmp_workspaces` E2E in `README.md`.
- [x] Add explicit regression coverage for `.frx` companion preservation in the userform path.
- [x] Re-run `tmp_workspaces` E2E verification for userform round-trip, including `.frm`/`.frx`.
- [x] Re-run `tmp_workspaces` E2E verification for `init` from an existing workbook after the hardening changes.

## Remaining

- [ ] Decide whether to broaden automated coverage beyond the current MVP hardening scope.

## Verification

- [x] Keep local verification entry point available via `task verify`.
- [x] Validate userform real workbook behavior with the `xlflow-tmp-workspace-e2e` skill.
- [x] Validate `init` real workbook behavior with the `xlflow-tmp-workspace-e2e` skill.
