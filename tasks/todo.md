# VBA Syntax Lint Todo

- [x] Add always-on lint findings for procedure boundary structure and line-continuation underscore whitespace.
- [x] Include the new syntax findings in push/run source-preflight blocking issues.
- [x] Add focused lint and CLI preflight regression tests.
- [x] Update CLI contract and README lint rule documentation.
- [x] Run `go test ./internal/lint ./internal/cli`, `go test ./...`, and `task lint`.

## Verification Notes

- `go test ./internal/lint ./internal/cli` passed.
- `go test ./...` passed.
- `task lint` passed.
- Excel COM E2E was not run because this change only affects static lint/source preflight logic and does not change workbook import/export, VBIDE automation, or macro execution behavior.

# xlflow Folder Structure Todo

# UserForm Phase 1 Warning Todo

- [x] Add shared UserForm detection/warning helpers in `internal/excel/scripts/common.ps1`.
- [x] Surface UserForm warnings/hints in `pull`, `push`, and `session`/`save`.
- [x] Add file-based inspect warnings from configured `src.forms` `.frm` detection without Excel COM.
- [x] Add focused Go and PowerShell regression coverage for warning helpers and inspect/source detection.
- [x] Run focused verification and `go test ./...`.
- [x] Update CLI/README/skill docs for Phase 1 warning behavior.

## Verification Notes

- `go test ./internal/cli ./internal/output ./internal/excel -run "TestCollectSourceUserFormNamesFindsRecursiveFrmFiles|TestInspectSourceUserFormMessagesReturnsWarningAndHint|TestWriteWithOptionsRendersInspectSnapshotMetadata"` passed.
- `go test ./internal/excel/scripts -run "TestGetXlflowSourceUserFormNamesFindsRecursiveFrmFiles|TestAddXlflowUserFormMessagesAddsDiscoveryAndStaleWarnings|TestPushScriptScopesSaveSessionWarningToSessionRuns|TestSessionStatusTreatsUnknownDirtyStateAsSaveRequired|TestPowerShellScriptsParse"` passed.
- `go test ./...` passed.

# UserForm Phase 2 List Forms Todo

- [x] Add `xlflow list forms` CLI command and session/keepalive flag plumbing.
- [x] Add Go bridge support plus top-level `forms` envelope/output rendering.
- [x] Add `internal/excel/scripts/list.ps1` for workbook UserForm discovery and folder-aware source path reporting.
- [x] Add focused Go and PowerShell regression coverage for args, parsing, and human output.
- [x] Update CLI contract, README files, and bundled xlflow skill guidance.
- [x] Run focused verification plus `go test ./...`.

## Verification Notes

- `go test ./internal/cli ./internal/output ./internal/excel -run "TestRootCommandIncludesListFormsCommand|TestRootCommandIncludesExcelCommandKeepaliveFlags|TestRootCommandIncludesSessionFlagsForWorkbookReaders|TestListFormsScriptArgsIncludeFolderAndSessionConfig|TestWriteWithOptionsRendersListFormsSummary"` passed.
- `go test ./internal/excel/scripts -run "TestPowerShellScriptsParse|TestListScriptValidatesActionBeforeWorkbookOpen|TestListScriptUsesFormComponentPathAndPortableRelativePaths"` passed.
- `go test ./...` passed.

# UserForm Phase 3 Inspect Form Todo

- [x] Add `xlflow inspect form` CLI command with basis, initializer, session, and keepalive flags.
- [x] Add Go bridge plumbing plus inspect payload/output support for form snapshots.
- [x] Add `inspect-form.ps1` and temporary VBA helper module support for runtime inspection.
- [x] Keep designer-only inspection macro-safe by inspecting VBIDE Designer state directly from PowerShell.
- [x] Add focused Go and PowerShell regression coverage.
- [x] Validate against `tmp_workspaces\user-form\build\Book.xlsm` with the existing `UserForm1`.
- [x] Update feature spec, CLI contract, and README files.

## Verification Notes

- `go test ./internal/cli ./internal/excel ./internal/output` passed.
- `go test ./internal/excel/scripts -run "TestPowerShellScriptsParse|TestInspectFormScriptValidatesBasisBeforeWorkbookOpen|TestInspectFormScriptUsesTemporaryHelperModuleAndWarnings"` passed.
- Runtime validation workspace: `C:\dev\go\xlflow\tmp_workspaces\user-form`.
- `go run C:\dev\go\xlflow\cmd\xlflow --json inspect form UserForm1 --runtime --initializer InitializeForm` returned `Order Entry Form`, width `308`, height `372`, `Alpha Stores`, `H001`, `Pencil`, `100`, and quantity `1`.
- `go run C:\dev\go\xlflow\cmd\xlflow --json inspect form UserForm1 --designer` passed.
- `go run C:\dev\go\xlflow\cmd\xlflow --json inspect form UserForm1 --both --initializer InitializeForm` passed.

# UserForm Phase 4 Snapshot Todo

- [x] Add `xlflow form snapshot <name> --out <path>` CLI command with session and keepalive support.
- [x] Reuse `InspectForm` designer output and add Go-side spec conversion plus JSON/YAML serialization.
- [x] Validate snapshot output extensions strictly against `.json`, `.yaml`, and `.yml`.
- [x] Add focused Go regression coverage for command wiring, argument validation, spec conversion, output rendering, and file writing.
- [x] Update CLI contract, README files, bundled skill guidance, and dependency/licence metadata for YAML support.
- [x] Validate `form snapshot` against `tmp_workspaces\user-form\build\Book.xlsm` for both JSON and YAML output.

## Verification Notes

- `go test ./internal/cli ./internal/excel ./internal/output -run "FormSnapshot|InspectForm|ExportImage"` passed.
- `go test ./internal/cli ./internal/excel ./internal/output ./internal/agentskill` passed.
- `go test ./internal/excel/scripts -run "TestPowerShellScriptsParse|TestInspectFormScriptValidatesBasisBeforeWorkbookOpen|TestInspectFormScriptUsesTemporaryHelperModuleAndWarnings"` passed.
- `powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\check-third-party-licences.ps1` passed.
- `go test ./...` passed.
- Validation workspace: `C:\dev\go\xlflow\tmp_workspaces\user-form`.
- `go run ..\..\cmd\xlflow --json form snapshot UserForm1 --out artifacts\UserForm1.form.json` returned `command=form snapshot`, `forms.name=UserForm1`, `forms.basis=designer`, `forms.control_count=14`, and `output.path=artifacts/UserForm1.form.json`.
- `go run ..\..\cmd\xlflow --json form snapshot UserForm1 --out artifacts\UserForm1.form.yaml` returned `command=form snapshot`, `forms.name=UserForm1`, `forms.basis=designer`, `forms.control_count=14`, and `output.path=artifacts/UserForm1.form.yaml`.
- Persisted JSON/YAML snapshots included `schemaVersion`, `kind`, `basis`, `coordinateSystem`, `form`, `controls`, and `warnings`, with camelCase spec fields such as `tabIndex` and `selectedIndex`.
- `go test ./internal/excel/scripts -run 'TestInspectFormScriptDesignerDoesNotRequireRunnableVBA|TestInspectFormScriptStrictDesignerReturnsConcreteControlTypes' -v` passed and confirmed direct `inspect form --designer` remains compile-tolerant while the strict snapshot/helper path returns concrete control types (`Label`, `TextBox`).

# UserForm Phase 5 Form Export Image Todo

- [x] Add `xlflow form export-image <name> --out <path.png>` CLI command with initializer, overwrite, session, and keepalive support.
- [x] Add Go-side option/path validation and PowerShell bridge plumbing for runtime UserForm image export.
- [x] Add `internal/excel/scripts/form-export-image.ps1` with temporary workbook copy execution, helper module injection, caption-token window lookup, and PNG capture/cleanup.
- [x] Extend JSON envelope and human output rendering for `form export-image` target/forms/output metadata and experimental/runtime warnings.
- [x] Update CLI contract, README files, bundled skill guidance, and UserForm hint text to point to `xlflow form export-image`.
- [x] Add focused Go and PowerShell regression coverage.
- [x] Run focused verification, full `go test ./...`, and Windows Excel COM validation against `tmp_workspaces\user-form`.

# UserForm Phase 6-7 Forms Package / Build Apply Todo

- [x] Extract UserForm spec parsing, validation, and snapshot serialization into `internal/excel/forms`.
- [x] Keep existing `form snapshot` / `inspect form` wiring working through the new forms package.
- [x] Add `xlflow form build` CLI command with `--overwrite`, `--session`, `--no-save`, and keepalive support.
- [x] Keep `xlflow form apply` implemented but hidden while the public replacement workflow moves to `form build --overwrite`.
- [x] Add Go-side spec loading/validation before Excel open and bridge payload serialization for build/apply.
- [x] Add `internal/excel/scripts/form-write.ps1` using the VBIDE Designer API for build/apply.
- [x] Extend human output rendering for `form build` / `form apply`.
- [x] Update CLI contract, README files, feature spec, and UserForm hint text for build/apply.
- [x] Run focused tests, full `go test ./...`, and Windows Excel COM validation for build/apply.

## Verification Notes

- `go test ./internal/cli ./internal/excel ./internal/output` passed.
- `go test ./internal/excel/forms` passed.
- `go test ./internal/excel/scripts -run "TestPowerShellScriptsParse|TestFormWriteScriptValidatesArgsBeforeWorkbookOpen|TestFormWriteScriptRejectsOverwriteWithNoSaveBeforeWorkbookOpen|TestFormWriteScriptUsesDesignerApiAndSessionSaveWarnings|TestFormWriteScriptDecodesSpecInWindowsPowerShell|TestInspectFormScriptUsesTemporaryHelperModuleAndWarnings|TestCommonScriptStrictDesignerFiltersControlsByParentName"` passed.
- `go test ./...` passed.
- Excel COM validation workspace: `C:\dev\go\xlflow\tmp_workspaces\form-build-apply-e2e`.
- `xlflow new --json`, `xlflow doctor --json`, `xlflow pull --json`, `xlflow lint --json`, `xlflow form build specs/form-build.json --json`, `xlflow list forms --json`, `xlflow inspect form DemoForm --json`, `xlflow form snapshot DemoForm --out specs/demo-snapshot.json --json`, `xlflow session start --json`, `xlflow form apply specs/form-apply.json --session --no-save --json`, `xlflow inspect form DemoForm --session --json`, `xlflow save --session --json`, `xlflow session stop --json`, `xlflow inspect form DemoForm --designer --json`, and `xlflow form export-image DemoForm --out specs/DemoForm.png --json` all completed. Public guidance now prefers `form build --overwrite` instead of `form apply`.
- Verified `form build` creates `DemoForm`, `list forms` reports `src/forms/DemoForm.frm`, `form snapshot` writes `specs/demo-snapshot.json`, and hidden `form apply --session --no-save` still returns `workbook.dirty=true` / `workbook.needs_save=true` / `session.save_required=true`. Public replacement workflow is rebuild via `form build --overwrite`.
- Follow-up COM validation workspace: `C:\dev\go\xlflow\tmp_workspaces\form-build-contract-e2e`.
- `xlflow new --json`, `xlflow doctor --json`, `xlflow pull --json`, `xlflow lint --json`, `xlflow form build src/forms/CustomerForm.form.json --json`, `xlflow inspect form CustomerForm --designer --json`, `xlflow form snapshot CustomerForm --out artifacts/CustomerForm.snapshot.json --json`, `xlflow form build src/forms/CustomerForm.overwrite.form.json --overwrite --json`, `xlflow inspect form CustomerForm --designer --json`, `xlflow form snapshot CustomerForm --out artifacts/CustomerForm.overwrite.snapshot.json --json`, and a final `xlflow pull --json` all completed after reinstalling `xlflow`.
- Verified `form build --overwrite` now succeeds after an intermediate workbook save, strict designer inspect no longer duplicates nested controls at the top level, and persisted snapshot specs now use flat `controls` entries with `id` / `parentId` / `zIndex`.
- Regression validation workspace: `C:\dev\go\xlflow\tmp_workspaces\form-overwrite-restore-e2e`.
- `xlflow new --json`, `xlflow doctor --json`, `xlflow form build src/forms/StableForm.form.json --json`, `xlflow inspect form StableForm --designer --json`, `xlflow form build src/forms/StableForm.bad.form.json --overwrite --json`, and `xlflow inspect form StableForm --designer --json` completed.
- Verified a failed `form build --overwrite` with an invalid runtime ProgID now restores the original UserForm and leaves the workbook with the pre-overwrite form still present.
- Session regression validation workspace: `C:\dev\go\xlflow\tmp_workspaces\form-overwrite-restore-session-e2e`.
- `xlflow new --json`, `xlflow doctor --json`, `xlflow session start --json`, `xlflow form build src/forms/SessionForm.form.json --session --json`, `xlflow inspect form SessionForm --designer --session --json`, `xlflow form build src/forms/SessionForm.bad.form.json --overwrite --session --json`, `xlflow inspect form SessionForm --designer --session --json`, and `xlflow session stop --json` completed.
- Verified the same failed overwrite path restores the original UserForm in the live session workbook as well; the session remained `dirty=false` / `save_required=false` after restoration.
- `form write` now emits explicit contract warnings for weak Designer-backed fields: `best_effort_form_size` for form-level `width` / `height`, and `best_effort_list_state` for design-time `ComboBox` / `ListBox` `list` / `selectedIndex`.
- Remaining known limitations from COM validation: form-level width/height still do not round-trip through the Designer API surface we can currently reach, and design-time `ComboBox` / `ListBox` item lists plus `selectedIndex` remain unreliable enough that build currently returns warnings and snapshots may come back with empty `list` and `selectedIndex=-1`.

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

# Explicit Workbook State Output Todo

- [x] Add shared `target`, `session`, `warnings`, and `hints` result shape support in the PowerShell bridge and Go envelopes.
- [x] Emit explicit target/session state from `session`, `push`, `pull`, `run`, `save`, `macros`, and `export-image`.
- [x] Keep `inspect` file-based while probing the matching recorded session for dirty/save-required warnings when available.
- [x] Extend human-readable output to show target and session state clearly, plus contextual warnings and hints.
- [x] Add regression coverage for JSON envelope fields, inspect state output, and PowerShell helper serialization.
- [x] Run focused verification plus full `go test ./...`.

## Verification Notes

- `go test ./internal/output ./internal/cli ./internal/excel` passed.
- `go test ./internal/excel/scripts` passed.
- `go test ./...` passed.

# xlflow Workbook Edit Commands Todo

- [x] Add `edit cell|range|rows|columns` CLI commands and flag validation.
- [x] Add Go-side edit option types, argument normalization, and bridge plumbing.
- [x] Add `edit.ps1` session-only Excel COM implementation with event-state restore.
- [x] Extend the JSON envelope and human renderer with top-level `edit` metadata.
- [x] Update CLI contract, README files, and any required workflow guidance.
- [x] Decide whether this policy change needs an ADR update and record the result (`docs/adr/ADR-0004-explicit-excel-session-mode.md` updated in this PR).
- [x] Add focused Go and PowerShell regression coverage.
- [x] Run focused `go test` targets, full `go test ./...`, and Windows Excel COM E2E verification.

## Verification Notes

- `go test ./...` passed.
- `task lint` passed.
- Workspace `C:\dev\go\xlflow\tmp_workspaces\edit-review-e2e`: `xlflow new`, `doctor`, `pull`, `lint`, `session start`, `push --fast --session --no-save`, `edit cell|range|rows|columns --session`, `run Main.Run --session`, `save --session`, `session stop`, `pull`, and final `lint` all passed.
- Excel COM workbook-state check after save/stop confirmed `A1="xlflow ok"`, `B2Formula="=1+2"`, `B2Value=3`, `C1Color=65280`, `Row1Height=24`, and `ColumnBWidth=22`.

# UserForm Code-Behind Sidecar Todo

- [x] Add shared PowerShell helpers for `src/forms/code/*.bas` sidecar discovery, export, and CodeModule reapplication.
- [x] Update `pull` to export UserForm code-behind sidecars and `push` to reapply them after `.frm` import.
- [x] Update `form build --overwrite` to preserve code-behind by preferring the sidecar and falling back to the pre-delete workbook form code.
- [x] Update CLI contract, README files, and bundled skill references for the `spec + code sidecar` source-of-truth model.
- [x] Add focused Go/PowerShell regression coverage and rerun lint/full tests.
- [x] Run Windows Excel COM E2E for UserForm build/pull/push/overwrite with code-behind preservation.

## Verification Notes

- `go test ./internal/excel/forms ./internal/excel/scripts ./internal/excel ./internal/agentskill` passed.
- `task lint` passed.
- `go test ./...` passed.
- Workspace `C:\dev\go\xlflow\tmp_workspaces\userform-codebehind-sidecar-e2e`: `xlflow new`, `doctor`, `pull`, `lint`, `form build src/forms/specs/CalendarPicker.yaml`, `pull`, `push`, `form build src/forms/specs/CalendarPicker.yaml --overwrite`, and final `pull` all passed.
- Excel COM workbook-state check after overwrite confirmed the rebuilt `CalendarPicker` form still contained code-behind version `B` even after deleting `src/forms/code/CalendarPicker.bas`, proving overwrite fallback preserved workbook code.
- Final `pull` recreated `src/forms/code/CalendarPicker.bas` with the preserved `B` code-behind.

# UserForm Code Source Mode Hardening Todo

- [x] Add `[userform].code_source = "frm" | "sidecar"` config, validation, and scaffold defaults (`new=sidecar`, `init=frm`).
- [x] Make `pull`, `push`, and `form build` mode-aware so `frm` mode ignores `src/forms/code` while `sidecar` mode exports and reapplies code-behind sidecars.
- [x] Synchronize tracked `.frm` embedded code from `src/forms/code/*.bas` before `push` and `form build` in `sidecar` mode so sidecar-only edits remain runnable.
- [x] Run `form build` through the same source preflight used by `push`/`run` when `sidecar` mode may inject VBA.
- [x] Update CLI contract, README files, and bundled skill references for mode-aware UserForm source-of-truth behavior.
- [x] Add focused Go and PowerShell regression coverage for config defaults, sidecar preflight, and `.frm` artifact synchronization.
- [x] Run Windows Excel COM E2E for both `new` (`sidecar`) and `init` (`frm`) workflows.

## Verification Notes

- `go test ./internal/config ./internal/project ./internal/excel/forms ./internal/cli ./internal/excel ./internal/excel/scripts ./internal/agentskill` passed.
- `task lint` passed.
- `go test ./...` passed.
- Workspace `C:\dev\go\xlflow\tmp_workspaces\userform-code-source-mode-e2e` (`new`, `code_source=sidecar`): `xlflow new`, `doctor`, `pull`, `lint`, `form build src/forms/specs/CalendarPicker.yaml`, `pull`, `push`, `pull`, `form build src/forms/specs/CalendarPicker.yaml --overwrite`, and final `pull` all passed.
- Sidecar-mode E2E confirmed `pull -> edit src/forms/code/CalendarPicker.bas -> push -> pull` preserved code-behind version `B`, and a manually diverged `src/forms/CalendarPicker.frm` was synchronized back to the authoritative sidecar before `form build --overwrite`.
- Workspace `C:\dev\go\xlflow\tmp_workspaces\userform-code-source-mode-init` (`init`, `code_source=frm`): `xlflow init <existing workbook>`, `pull`, `.frm` code edit, `push`, `pull`, `form snapshot`, `form build --overwrite`, and final `pull` all passed.
- FRM-mode E2E confirmed `pull` did not create `src/forms/code/*.bas`, `push` preserved `.frm`-embedded code version `FRM2`, and `form build --overwrite` kept that embedded code intact.
- Workspace `C:\dev\go\xlflow\tmp_workspaces\userform-preflight-scope-e2e`: `xlflow new`, `doctor`, `pull`, `lint`, then `form build src/forms/specs/UserForm1.yaml` in a sidecar repo where unrelated `src/forms/UserForm2.frm` intentionally contained stale analyzer-breaking code. Build for `UserForm1` still succeeded, confirming form-build preflight now scopes UserForm source checks to the target form instead of unrelated generated `.frm` artifacts.
