# xlflow Runtime Debugging Hardening Todo

## Skill provider target correction

- [x] Remove `copilot` from provider defaults and CLI help.
- [x] Add regression coverage that `--agent copilot` is unsupported.
- [x] Update README, CLI contract, and ADR provider target text.
- [x] Run focused Go tests for skill install behavior.

## Phase 1: Persist trace injection

- [x] Update the CLI contract for `trace inject` source persistence.
- [x] Pass configured source module directory information from Go to `scripts/trace.ps1`.
- [x] Add a reusable helper that writes bundled `XlflowTrace` source as UTF-8 without BOM.
- [x] Make `trace inject` write `src/modules/XlflowTrace.bas` when running against the configured project workbook.
- [x] Keep explicit standalone workbook injection working without project configuration.
- [x] Return JSON source metadata such as `source.path` and `source.updated`.
- [x] Add tests proving `trace inject` followed by `push` preserves `XlflowTrace`.
- [x] Update README and bundled skill trace guidance.

## Phase 2: Improve run failure diagnostics

- [x] Define stable `run` failure phases in `docs/specs/cli-contract.md`.
- [x] Track the current phase in `scripts/run.ps1`.
- [x] Include the phase in structured JSON errors.
- [x] Split missing or invalid macro target failures from ordinary user-code `macro_failed` when Excel exposes enough information.
- [x] Improve plain-text failure messages without making successful output noisy.
- [x] Add tests for phase metadata and macro target failure output.
- [x] Update the bundled skill to explain `error.phase` and `macro_not_found` handling.

## Phase 3: Add macro entrypoint discovery

- [x] Choose the public command shape: prefer `xlflow macros --json` unless implementation constraints favor `xlflow run --list --json`.
- [x] Specify the discovery JSON schema.
- [x] Implement workbook macro discovery without executing user code.
- [x] Include module name, procedure name, fully qualified name, and argument information when available.
- [x] Add tests for command wiring and PowerShell discovery parsing.
- [x] Document the workflow: discover entrypoints before guessing a `run` target.
- [x] Update the bundled skill to use macro discovery before guessing a `run` target.

## Phase 4: Detect automation-hostile VBA patterns

- [x] Add lint rules for `Application.GetOpenFilename`, `Application.FileDialog`, `InputBox`, and modal `MsgBox`.
- [x] Assign stable lint codes.
- [x] Add fixture coverage for each pattern.
- [x] Update lint documentation with recommended CLI-friendly alternatives.
- [x] Update the bundled skill to discourage UI prompts in agent-run macros.

## Phase 5: Improve empty trace guidance

- [x] Add a structured hint or log when `run --trace` fails with zero trace events.
- [x] Ensure failures after partial trace logging still return existing events.
- [x] Update the bundled skill with the zero-event interpretation.
- [x] Add tests for empty-trace failure output.

## Final verification

- [x] Run `go test ./...` with a long timeout.
- [x] Run `task verify`.
- [x] Run the `xlflow-tmp-workspace-e2e` workflow when Excel/VBIDE access is available.
- [x] Review docs for consistency across `README.md`, `docs/specs`, and bundled skill text.
