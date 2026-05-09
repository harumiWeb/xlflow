# xlflow Folder Structure Todo

- [x] Add `[vba]` config defaults and validation.
- [x] Make `pull` folder-aware and clear stale recursive exports.
- [x] Make `push` import recursive source trees and preserve nested `.frm`/`.frx` companions.
- [x] Rewrite temporary import annotations from filesystem paths in `update` mode.
- [x] Detect duplicate VBA module names before Excel import.
- [x] Add focused Go and PowerShell regression coverage for folder-aware behavior.
- [ ] Update broader docs and examples if folder mode UX changes further.

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

- [x] Add `version --verbose` payload and human output.
- [x] Auto-reuse matching sessions for `pull`, `push`, `macros`, `run`, `test`, `trace`, and `save`.
- [x] Promote unsaved live-session state to structured `needs_save` metadata and stronger human output.
- [x] Report dirty/save-required state from `session status`.
- [x] Keep `push` save-by-default semantics while preserving `--no-save` as the session opt-out.
- [x] Update the bundled skill for auto session reuse and `run` macro omission fallback.
- [x] Add focused Go and PowerShell regression coverage.

## Verification Notes

- `go test ./internal/cli ./internal/output ./internal/excel -run "TestVersionCommandVerboseIncludesExecutableAndFeatures|TestWriteWithOptionsRendersVersionVerboseDetails|TestWriteWithOptionsRendersSessionOnlyPushResult|TestWriteWithOptionsRendersSessionStatusSaveRequirement|TestBuildRunOptions|TestBuildPushOptions|TestRootCommandIncludesSessionFlagsForWorkbookReaders|TestTraceScriptArgsPassSessionFlag|TestBuildRunScriptArgsPassesFastDirectAndSession"` passed.
- Auto-reuse is implemented in `internal/excel/scripts/common.ps1` via `Open-XlflowWorkbookForCommand`, `Test-XlflowSessionMetadataMatchesWorkbook`, and session-mode result metadata.
- Human output and JSON now distinguish `session_mode=explicit|auto|managed|none` and surface `workbook.needs_save`, `workbook.dirty`, `session.needs_save`, and `session.dirty`.

# Release Artifact Trust Hardening Todo

- [x] Add `checksum.algorithm: sha256` to `.goreleaser.yaml` while keeping `checksums.txt` stable.
- [x] Enable GoReleaser SBOM generation for archive artifacts.
- [x] Extend `release.yml` permissions for artifact attestation.
- [x] Install Syft in the release workflow before GoReleaser runs.
- [x] Attest release archives via `dist/checksums.txt`.
- [x] Attest `dist/checksums.txt` itself.
- [x] Attest generated SBOM artifacts.
- [x] Document SHA256 verification in `README.md`.
- [x] Document SHA256 verification in `README.ja.md`.
- [x] Document attestation verification and the Authenticode non-claim in both READMEs.
- [x] Run `goreleaser check`.
- [x] Run `goreleaser release --snapshot --clean --skip=publish`.
- [x] Run `go test ./...`.
- [x] Validate workflow syntax with `actionlint` if available.

## Verification Notes

- `goreleaser check` passed.
- `actionlint .github/workflows/release.yml` passed.
- `go test ./...` passed.
- `goreleaser release --snapshot --clean --skip=publish` passed after installing a working `syft` binary.
- Snapshot artifacts included `dist/checksums.txt`, `dist/xlflow_windows_x86_64.zip`, and `dist/xlflow_windows_x86_64.zip.sbom.json`.
- `dist/checksums.txt` contained SHA256 entries for both the release ZIP and the generated SBOM file.

# Security and Licence Automation Todo

- [x] Add a CI `govulncheck` job on Windows using the Go toolchain from `go.mod`.
- [x] Add a repo-local licence inventory checker for `THIRD_PARTY_LICENCES.md` against `go list -deps ./cmd/xlflow`.
- [x] Expose the same checks through `task verify:security`.

# xlflow Style-Aware Inspect Todo

- [x] Add `--include-style` plumbing to `inspect range` and `inspect used-range`.
- [x] Extend file-based inspect snapshots with target metadata and style-aware range payloads.
- [x] Add focused inspect and CLI regression coverage for styled ranges, merged cells, row/column metadata, and compatibility without `--include-style`.
- [x] Update CLI contract, README files, and bundled xlflow skill guidance for style-aware inspect.
- [x] Run focused tests plus `go test ./...` and record results.

## Verification Notes

- `go test ./internal/inspect ./internal/cli ./internal/output` passed.
- `go test ./...` passed.

# xlflow Range Image Export Todo

- [x] Add `export-image` CLI command and flag validation.
- [x] Add Go-side export-image option/path resolution and PowerShell bridge plumbing.
- [x] Add `export-image.ps1` Excel COM implementation with temporary chart cleanup.
- [x] Extend output envelope and human rendering for export target/output/warnings metadata.
- [x] Update CLI contract, README files, bundled skill guidance, and ADR session-reuse note.
- [x] Add focused Go and PowerShell regression coverage.
- [x] Run focused verification and full `go test ./...`.
- [x] Run Excel COM-backed end-to-end export verification in a disposable workbook workspace.

## Verification Notes

- `go test ./internal/cli ./internal/excel ./internal/output` passed.
- `go test ./internal/excel/scripts -run TestPowerShellScriptsParse` passed.
- `go test ./...` passed.
- Excel COM-backed end-to-end export verification was run in a disposable workbook workspace during this pass.
