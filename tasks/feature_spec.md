# xlflow MVP Quality Hardening Spec

## Goal

Harden post-MVP quality checks by locking in verified module-round-trip behavior, regression protection, and a clear local verification entry point before the next real Excel COM verification pass.

## Active Slice

Current hardening slice status:

- [x] Task 1: `docs/specs/cli-contract.md` and `README.md` aligned with verified behavior.
- [x] Task 2: regression tests plus `scripts/common.ps1` hardening for document-module normalization, including malformed-header protection.
- [x] Task 3: `Taskfile.yml` `verify` target added and README clarifies local verify vs `tmp_workspaces` E2E.
- [x] Follow-up: explicit `.frx` companion regression coverage added.
- [x] Follow-up: rerun the real workbook verification path (`xlflow-tmp-workspace-e2e`) for userforms and `init` after the hardening changes.

## Verification Focus

1. keep local fast checks (`task verify`) as the routine gate
2. keep `.frx` companion preservation covered by the automated script suite
3. keep `tmp_workspaces` E2E for userform and `init` paths as the real Excel COM proof
