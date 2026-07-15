# ADR-0005: Runtime Execution Mode Injection for VBA

## Status

`accepted`

## Background

xlflow already distinguishes human-assisted and unattended execution at the CLI layer through modes such as `run --headless`, `run --interactive`, and `test`. However, workbook VBA code cannot currently read that execution context directly. Projects that need different behavior for interactive Excel usage, unattended agent runs, CI, or test execution would otherwise have to rely on brittle techniques such as process inspection, parent-process guesses, or UI heuristics.

Alternative approaches considered:

- Keep runtime context only in the CLI and require every project to invent its own VBA-side detection.
- Detect automation context from Excel visibility, window state, or process ancestry.
- Inject workbook-scoped runtime state before user VBA starts and expose it through a scaffolded VBA helper.
- Use only environment variables and do not write any workbook-scoped state.

The runtime contract needs to be explicit, deterministic, and available to ordinary VBA code without introducing a second automation backend.

## Decision

xlflow injects workbook-scoped reserved names before `run` and `test` invoke user VBA, restores those names afterward, and scaffolds the standard module `src/modules/Xlflow/XlflowRuntime.bas` in new projects so VBA can query the resolved execution mode.

The primary signal is workbook-scoped state, currently including `__XLFLOW_MODE__` and a compatibility version marker. `XlflowRuntime.bas` reads that workbook-scoped state first and uses `Environ$("XLFLOW_MODE")` only as a secondary fallback for older projects or wrapper-driven runs that adopt the helper manually.

Phase 1 runtime modes are:

- `interactive`
- `headless`
- `ci`
- `agent`
- `test`

Resolution rules in Phase 1 are:

- `run --headless` resolves to `headless`.
- `run --interactive` resolves to `interactive`.
- `test` resolves to `test`.
- Plain `run` defaults to `interactive` unless the xlflow process environment sets `XLFLOW_MODE=interactive|headless|ci|agent|test`.
- Unknown `XLFLOW_MODE` values are ignored instead of guessed.

The public VBA helper surface is module-qualified rather than pseudo-namespaced:

- `XlflowRuntime.Mode()`
- `XlflowRuntime.ModeName()`
- `XlflowRuntime.IsInteractive()`
- `XlflowRuntime.IsHeadless()`
- `XlflowRuntime.IsCI()`
- `XlflowRuntime.IsAgent()`
- `XlflowRuntime.IsTest()`

Existing projects are not auto-migrated in Phase 1. `xlflow new` scaffolds `XlflowRuntime.bas` under `src/modules/Xlflow/`; imported projects can add it manually when they opt into runtime-aware branching.

## Consequences

Positive consequences:

- VBA code can branch on an explicit execution contract without inferring xlflow behavior from UI state or process ancestry.
- Agents, tests, and CI can share one deterministic runtime vocabulary.
- The runtime contract stays inside the existing Go CLI plus PowerShell bridge architecture.
- The helper remains usable in ordinary VBA code and is source-controlled like the rest of the project modules.

Negative consequences:

- xlflow must carefully clean up temporary runtime markers so they do not leak into workbook state or distort save-required reporting.
- Workbook-scoped markers require Excel/VBIDE-backed commands and therefore remain Windows/Excel-specific like the rest of the automation bridge.
- Existing projects do not gain the helper automatically; documentation must explain the manual adoption path.
- Runtime mode injection alone does not solve GUI-boundary design. `run --headless` still uses project-wide GUI preflight and does not become call-graph-aware.

## Rationale

- Tests: focused Go tests for runtime mode resolution and script argument propagation, PowerShell script parse/behavior tests for runtime marker injection and cleanup, plus Windows Excel COM validation that proves mode values reach VBA.
- Code: `internal/cli`, `internal/excel`, `internal/excel/scripts/run.ps1`, `internal/excel/scripts/test.ps1`, `internal/project/scaffold.go`.
- Related specs: `docs/specs/cli-contract.md`, `docs/specs/runtime-debugging.md`, `docs/design.md`.

## Supersedes

- None

## Superseded by

- None
