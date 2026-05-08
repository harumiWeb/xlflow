# MVP Quality Hardening

## Goal

Stabilize the `xlflow` MVP so future feature work can rely on repeatable Excel COM validation, consistent CLI contracts, and regression coverage for workbook round-trips.

## Why Now

The MVP command set works for the core flow, but recent verification exposed failure modes that unit tests alone did not catch:

- document module exports needed source normalization before re-import
- workbook save behavior had to be verified through Excel COM, not only CLI success output
- class and userform round-trips required real workbook fixtures and `.frx` companion artifacts

Before adding new product features, the repository should make these behaviors easier to verify and harder to regress.

## Scope

This phase covers four workstreams:

1. `tmp_workspaces`-based end-to-end verification
2. regression coverage for Excel VBA component round-trips
3. CLI contract hardening for machine-readable outputs
4. contributor workflow improvements so agents and humans run the same verification path

This phase does not add new user-facing commands.

## Required Outcomes

### 1. E2E Coverage Baseline

The repository must define a standard verification baseline that can be rerun from scratch under `tmp_workspaces`.

The baseline should cover:

- blank workbook scaffold:
  - `new`
  - `doctor`
  - `pull`
  - `lint`
- standard module macro round-trip:
  - `push`
  - `run`
  - `pull`
  - `lint`
- class module round-trip
- userform round-trip, including `.frm` and `.frx`
- `init` from an existing workbook

### 2. Regression Tests

Automated regression coverage should protect the highest-risk script behaviors:

- document module normalization
- import/export handling for `.bas`, `.cls`, `.frm`
- preservation of `.frx` companion artifacts for userforms
- stable JSON envelope fields and failure codes

The automated suite does not need to fully replace Excel COM E2E checks, but it must catch known transform and contract failures before manual verification.

### 3. CLI Contract Stability

`docs/specs/cli-contract.md` should remain the source of truth for:

- supported commands
- JSON envelope fields
- command-specific result payloads
- exit-code behavior

If implementation diverges from this contract, either the code or the spec must be updated in the same change.

### 4. Workflow Consistency

Repository tooling and guidance should direct contributors toward the same quality bar:

- use the repo-local `xlflow-tmp-workspace-e2e` skill for real workbook validation
- keep `tasks/lessons.md` in sync with new Excel/VBIDE failure patterns
- provide one obvious command or short sequence for local post-change verification
- treat Windows + Excel real-workbook E2E as a manual release gate for Excel COM changes, not as a normal CI substitute

## Non-Goals

- redesigning the CLI surface
- changing the workbook/project configuration format
- adding diff/merge UX or advanced VBA editing features
- supporting non-Windows Excel environments

## Deliverables

- updated tests for script and contract regressions
- updated docs/specs where behavior is clarified
- contributor-facing verification guidance or wrapper commands
- repeatable manual verification evidence using `tmp_workspaces`
- release-preflight guidance that points agents and humans at the same `xlflow-tmp-workspace-e2e` workflow

## Exit Criteria

This phase is complete when:

- blank workbook, module, class, form, and `init` verification paths have a documented and repeatable flow
- automated tests cover the known document-module and form/class round-trip regressions
- contributor guidance points to one standard E2E verification path
- the current implementation and `docs/specs/cli-contract.md` agree on the supported MVP behavior
