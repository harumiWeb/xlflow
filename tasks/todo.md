# xlflow Runtime Debugging Hardening Todo

## Phase 1: Persist trace injection

- [x] Update the CLI contract for `trace inject` source persistence.
- [ ] Pass configured source module directory information from Go to `scripts/trace.ps1`.
- [ ] Add a reusable helper that writes bundled `XlflowTrace` source as UTF-8 without BOM.
- [ ] Make `trace inject` write `src/modules/XlflowTrace.bas` when running against the configured project workbook.
- [ ] Keep explicit standalone workbook injection working without project configuration.
- [ ] Return JSON source metadata such as `source.path` and `source.updated`.
- [ ] Add tests proving `trace inject` followed by `push` preserves `XlflowTrace`.
- [ ] Update README and bundled skill trace guidance.

## Phase 2: Improve run failure diagnostics

- [x] Define stable `run` failure phases in `docs/specs/cli-contract.md`.
- [ ] Track the current phase in `scripts/run.ps1`.
- [ ] Include the phase in structured JSON errors.
- [ ] Split missing or invalid macro target failures from ordinary user-code `macro_failed` when Excel exposes enough information.
- [ ] Improve plain-text failure messages without making successful output noisy.
- [ ] Add tests for phase metadata and macro target failure output.
- [ ] Update the bundled skill to explain `error.phase` and `macro_not_found` handling.

## Phase 3: Add macro entrypoint discovery

- [ ] Choose the public command shape: prefer `xlflow macros --json` unless implementation constraints favor `xlflow run --list --json`.
- [x] Specify the discovery JSON schema.
- [ ] Implement workbook macro discovery without executing user code.
- [ ] Include module name, procedure name, fully qualified name, and argument information when available.
- [ ] Add tests for command wiring and PowerShell discovery parsing.
- [ ] Document the workflow: discover entrypoints before guessing a `run` target.
- [ ] Update the bundled skill to use macro discovery before guessing a `run` target.

## Phase 4: Detect automation-hostile VBA patterns

- [ ] Add lint rules for `Application.GetOpenFilename`, `Application.FileDialog`, `InputBox`, and modal `MsgBox`.
- [ ] Assign stable lint codes.
- [ ] Add fixture coverage for each pattern.
- [ ] Update lint documentation with recommended CLI-friendly alternatives.
- [ ] Update the bundled skill to discourage UI prompts in agent-run macros.

## Phase 5: Improve empty trace guidance

- [ ] Add a structured hint or log when `run --trace` fails with zero trace events.
- [ ] Ensure failures after partial trace logging still return existing events.
- [ ] Update the bundled skill with the zero-event interpretation.
- [ ] Add tests for empty-trace failure output.

## Final verification

- [ ] Run `go test ./...` with a long timeout.
- [ ] Run `task verify`.
- [ ] Run the `xlflow-tmp-workspace-e2e` workflow when Excel/VBIDE access is available.
- [ ] Review docs for consistency across `README.md`, `docs/specs`, and bundled skill text.
