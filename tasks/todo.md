# xlflow Performance Mode Todo

- [x] Add push fast flags and Go option validation.
- [x] Add push source fingerprint state and no-op changed-only skip.
- [x] Add run direct/fast/session flags and Go option validation.
- [x] Add PowerShell direct run path.
- [x] Add explicit Excel session commands and session attach support.
- [x] Add persistent runner module commands.
- [x] Update CLI contract, README, bundled skill, and ADR.
- [x] Add focused Go tests for CLI/script argument plumbing.
- [x] Run focused Go unit tests and PowerShell parse tests.
- [x] Run full Excel COM-backed script/e2e tests in an environment where they complete within the expected long timeout.

# xlflow Diagnostic Run Todo

- [x] Add diagnostic run CLI/spec plumbing.
- [x] Add PowerShell VBE compile dialog watcher and selection diagnostics.
- [x] Return structured compile diagnostics from `run.ps1`.
- [x] Polish runtime diagnostic shape and human output.
- [x] Update CLI/runtime docs, README files, bundled skill, and ADR note.
- [x] Add focused Go and PowerShell tests.
- [x] Run focused verification and document Excel COM-backed test results.

## Verification Notes

- `go test ./internal/cli ./internal/excel ./internal/output` passed.
- Focused `go test ./scripts -run "TestPowerShellScriptsParse|TestRunScriptAcceptsDiagnosticParameter|TestRunScriptRejectsDirectDiagnosticBeforeOpeningWorkbook|TestVBESelectionDiagnosticHandlesMissingPane|TestRunHarnessCodeConfiguresTraceBeforeMacro"` passed.
- `go test ./...` passed.
- `go test ./scripts -count=1` passed in 232.613s.
- E2E workspace: `C:\dev\go\xlflow\tmp_workspaces\diagnostic-run-e2e`.
- `xlflow new --json`, `xlflow doctor --json`, `xlflow pull --json`, `xlflow lint --json`, and `xlflow push --json` passed in the E2E workspace.
- `xlflow run Main.Run --diagnostic --json` returned `vba_compile_failed` with `phase=compile_vba`, `source=Main`, `line=4`, localized dialog text, VBE selection, and nearby code for an intentional `Option Explicit` compile failure.
- A process-window check found no remaining top-level `Microsoft Visual Basic for Applications` dialog after the diagnostic run.

# Manual Regression Follow-up

- [x] Fixed `--fast` run validation accidentally treating `Diagnostic=false` as enabled because a PowerShell boolean expression omitted function-call parentheses.
- [x] Fixed compile dialog watcher timing so delayed VBE dialogs are still captured after `Execute()` returns.
- [x] Hardened `push.ps1` temporary import path construction by resolving the backup root to an absolute path before building `.xlflow/tmp/import/<timestamp>`.
- [x] Set Excel `UserControl = true` on `session start` so COM-created visible sessions are less likely to close when the bridge process exits.
- [x] Fixed human output so diagnostic `message` renders once as either a scalar line or an array block, not both.
- [x] E2E workspace: `C:\dev\go\xlflow\tmp_workspaces\manual-regression-e2e`.
- [x] Verified `new`, `doctor`, `pull`, `lint`, `push`, `macros`, `session start`, delayed `session status`, `push --fast --session --no-save`, harness `run --session`, direct `run --fast --session`, `save --session`, `session stop`, and workbook cell state.
- [x] Re-verified diagnostic compile failure in `C:\dev\go\xlflow\tmp_workspaces\diagnostic-run-e2e`; no VBE dialog remained afterward.

# xlflow Session-Aware Defaults Todo

- [ ] Add `version --verbose` payload and human output.
- [ ] Auto-reuse matching sessions for `pull`, `push`, `macros`, `run`, `test`, `trace`, and `save`.
- [ ] Promote unsaved live-session state to structured `needs_save` metadata and stronger human output.
- [ ] Report dirty/save-required state from `session status`.
- [ ] Keep `push` save-by-default semantics while preserving `--no-save` as the session opt-out.
- [ ] Update the bundled skill for auto session reuse and `run` macro omission fallback.
- [ ] Add focused Go and PowerShell regression coverage.
