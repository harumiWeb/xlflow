# Issue #949: type-backed Set Quick Fixes

- [x] Reproduce PR #950's indexed-property binding finding and isolate accessor-local candidates by declaration range.
- [x] Verify accessor parameters/locals and preserve duplicate-binding rejection; prepare the review correction for push.

PR #950 follow-up: focused accessor regressions failed before the scope fix and
passed afterward, including snapshot/live paths and rejected ambiguous/Variant
bindings. Complete intel/LSP suites pass (75.518s / 163.472s), scoped lint reports
zero issues, and docs checks pass. A separate read-only review of this bounded
delta found no confirmed defects; it did not independently execute the tests.
This source-declaration identity correction does not change diagnostic or VBE
type semantics, so the established oracle evidence is retained without rerunning
Excel. Native editor application/Undo and new remote CI results remain separate.

- [x] Confirm merged #948 and stage-three scope; create #949 before implementation.
- [x] Share selected-line assignment diagnostics and add conservative Set corrections.
- [x] Cover type/binding ambiguity, source preservation, diagnostic association and discovery costs.
- [x] Verify VBE controls/pairs, affected packages, docs and formatting/lint.
- [x] Commit and complete independent review.

All five affected package suites passed (lint, intel, LSP, registry and oracle).
Focused intel regressions pass after filtering suppression once per request;
the final complete intel suite also passed (66.867s). Scoped golangci-lint reports
zero issues; docs and diff checks pass. Deterministic discovery coverage checks
that a non-assignment selection builds no type index, a positive selection
builds the index, and neither builds ProcedureIR/CFG across unrelated procedures.
No whole-analyzer performance comparison or native VS Code apply/Undo is claimed.

Fresh sequential VBE runs passed the known accept/reject controls,
`quickfix-set-object-before`/`quickfix-set-object-after`, `set-scalar-target`,
`scalar-assignment`, and `set-object-target`, all with compile invocation and
cleanup confirmed. New object fixtures were explicitly promoted as policy;
VB037 retains the existing rejected fixture and accepted negative controls.
Initial attempts failed fixture metadata validation and then rejected a mixed
promotion batch containing asserted fixtures; they promoted no evidence. Full
fixture schema validation and separate new/established runs resolved this.
Runtime equivalence and remote CI remain unverified.

Independent Pass 1 reviewed exact commit `2026d126` against `96471ade` and
found no confirmed in-scope defects; focused intel/LSP tests passed independently.
No fix batch or second review pass is required. The report and coordinator
disposition are retained under `tmp_workspaces/final-review/issue949-pass1/`.
The delivery was acknowledged and the worker released; all four dedicated
review tabs were closed and the temporary worktree was removed, with absence
confirmed through the terminal/worktree listings.

# Issue #947: local diagnostic Quick Fixes

- [x] Confirm merged #946, existing diagnostic predicates, and stage-two scope; create #947 before implementation.
- [x] Add conservative VB080/VB086 edits and safe VB001 insertion using shared diagnostic association.
- [x] Add source-preservation, suppression, EOL/Unicode, header, idempotency, and no-IR regressions.
- [x] Verify VBE evidence, affected packages and registry/docs; complete lint and formatting through the commit checks.
- [x] Commit the implementation and complete independent review.

Validation: all five affected package suites pass (lint, intel, LSP, rule registry,
oracle), and generated documentation checks pass. The local VBE run passed
`known-compile-accept`, `known-compile-reject`, `quickfix-local-before`, and
`quickfix-local-after`, with compile invocation and Excel cleanup confirmed.
The focused pair was explicitly promoted as policy evidence. Native VS Code
application/Undo, runtime equivalence and remote CI are not established by
these checks. Independent Pass 1 reviewed exact commit `42130ec1` against
`5771a62e` and found no confirmed in-scope defects; focused intel/lint tests
also passed independently. The report is retained under
`tmp_workspaces/final-review/issue947-pass1/report.md`. The delivery was
acknowledged, the externally created review terminal was closed after release,
and the watcher/setup/shell tabs and temporary review worktree were removed.

# PR #946 security CI follow-up

- [x] Confirm the failing security job and the official GO-2026-6809 fix.
- [x] Raise the Go patch baseline from 1.26.9 to 1.26.10; preserve CI scanning.
- [x] Verify the full vulnerability scan and licence inventory with the patched toolchain.

The patched toolchain reports zero affecting vulnerabilities; licence inventory
matches all 39 CLI modules. The full Go suite is also running during preparation
of this correction. Its result and the subsequent remote checks belong in the
existing PR #946 validation record, separately from the confirmed local scan.

# Issue #945: shared Quick Fix edits and VB084

- [x] Record the shared edit contract and VB084 applicability in ADR/specs.
- [x] Migrate VB044 and documentation actions to multiple edits with diagnostic association.
- [x] Add opt-in VB084 correction with suppression, source-preservation, and no-IR coverage.
- [x] Verify package tests, generated docs, formatting/lint, and self-review the final diff.
- [x] Complete independent final review and resolve confirmed in-scope findings.

Pass 1 disposition: the reported VB044 diagnostic-range mismatch is unsupported.
`issueAt` retains the full range internally but intentionally omits public end
positions; both published and associated diagnostics use the same one-unit
range. Added direct full-diagnostic equality and full-literal edit coverage across
LF/CRLF and Unicode literals, without changing production behavior. No confirmed
in-scope defect remains, so no second full review pass is required.
The complete intel package suite passes again after this regression coverage
(`40.823s`). The reviewer report and coordinator disposition are preserved under
`tmp_workspaces/final-review/issue945-pass1`; the review-only tabs and worktree
have been removed.

Validation: all five affected Go packages pass their complete suites; scoped
golangci-lint reports zero issues, and docs/format/diff checks pass. VBE controls
`known-compile-accept`/`known-compile-reject` and focused cases
`quickfix-step-one-before`/`quickfix-step-one-after` pass with confirmed cleanup;
both focused cases were explicitly promoted as maintainability observations.
Runtime equivalence, native VS Code application/Undo, and remote CI are not
independently verified by these checks.
The full pre-commit lint additionally required explicit returns after existing
nil/Fatal guards in eight tests; those changes do not affect production code.

# Issue #942: code-only Designer preservation

- [x] Confirm the existing equal-caption repair contract and define an opt-in preservation mode.
- [x] Validate FormSpecs while rejecting actual Designer changes; preserve normal authoring behavior.
- [x] Add CLI/file-push integration with safe changed-only handling and rejection coverage.
- [x] Run affected Windows tests and the commit-pinned Linux suite with exact Designer/reference comparison.
- [x] Review and document the verified behavior, then prepare the upstream PR.

Validation: implementation `1a0a94b3` passed the affected Windows pack/compiler,
file-push and full CLI tests (CLI: 185.827s), scoped golangci-lint, documentation
and staged formatting checks. The pinned Docker run passed all 12 related Go
packages, six helper tests, 29 Python tests and all six fixture groups
(`suite-g27iy6ti`, 83.464s, 92 CLI calls). All 13 reference/Designer comparisons
had zero differences, including Frame no-op and edited pack. Frame/stdTable
file push safely rejected existing lint errors; 19 expected rejection records
passed. Independent review findings in the added test fixtures were corrected
and rechecked with no remaining findings. Excel/VBE compile and runtime were
not performed. Remote CI results are separate from these local checks.

# Issue #939: preserve document-module source round trips

- [x] Add focused regressions for no-op and edited document source, unchanged saved member attributes, and rejected module headers.
- [x] Fix the document normalization inverse while retaining the existing export-side terminal-LF rule; retain only exact saved member-attribute statements in template pack.
- [x] Verify pack, file-pull, file-push, and CLI packages on Windows and in the Linux builder without changing analyzer or VBA semantics.
- [x] Update the pack source contract and Unreleased changelog, review the scoped diff, and commit the validated result locally.
- [x] Run the commit-pinned Phase1 suite in a standalone clone and verify the existing Frame lint rejection still preserves the source copy. All six groups passed against implementation `51420159712b79cccc4877ec523c3fbfe9144aae`; existing Frame/stdTable lint failures safely rejected file push.

# Issue #937: container host detection

- [x] Reproduce Docker/Podman marker precedence with focused host-detection tests.
- [x] Exclude recognized containers from WSL delegation without changing ordinary WSL safety gates.
- [x] Document the detection boundary and verify Windows/Linux regression suites.
- [x] Validate the saved-workbook round trip in Docker Desktop and record remaining limits.

Validation: implementation commit `45df6530` passes Windows wsl/filepush/cli
tests and all 11 related Linux packages. Scoped golangci-lint, docs and staged
format checks pass. Docker Desktop no-op/edit template pack, file push,
re-pull, saved-file inspect/diff and rejection controls pass; non-VBA ZIP
parts and the original workbook are preserved. Source and static-diagnostic
comparisons pass (3 existing warnings, no new findings). VBE compile/runtime,
ordinary WSL with a live Windows Excel session, and remote CI were not run.
The read-only independent review found no confirmed actionable findings.

# PR #935 review follow-up

- [x] Compare retained TabStrip observations with the authored/observed validation contract.
- [x] Stop bounding historical selectedIndex by the current authored tabs.
- [x] Review the native navigation coverage observation and record its limits.
- [x] Run affected lint/docs checks and prepare the fix and English PR description.

Disposition: the historical TabStrip selection finding is valid. Captured
selection now uses only captured tabs for upper-bound validation; the existing
YAML/JSON edit test expectations reflect successful removal and preserved
observations. Native problem navigation remains a coverage limitation rather
than a confirmed implementation defect. Scoped lint, docs, formatting and diff
checks passed; tests and Excel were not rerun for this follow-up.

# Agent skill FormSpec authoring guidance

- [x] Validate Pass 1 findings against compiler and normalization behavior.
- [x] Correct CheckBox state guidance and clarify legacy nested controls in generated references.
- [x] Verify review fixes and complete fresh Pass 2 independent review.

- [x] Separate workflow, property guidance, generated contract and complete examples.
- [x] Document field types, enums, applicability, edit semantics and missing-property limits.
- [x] Generate property metadata from the shared contract and detect stale documentation.
- [x] Review links, formatting and generated output against the implementation.

Validation: documentation checks/build, scoped formatting, generator lint and
diff checks pass. Generated metadata is bundled with the skill and checked
without rewriting files. Read-only review identified legacy list-persistence
advice; guidance and metadata descriptions now distinguish ComboBox string
values from rejected list/selection and ListBox text/value edits. No runtime
semantics changed. Independent Pass 2 reviewed commit `22d38be1` and passed
without findings; spec/compiler/agentskill tests, docs checks and formatting
passed. Review reports are retained under `tmp_workspaces/final-review`, and
both temporary review environments were removed. Excel and remote Linux CI
were not run for this update.

# UserForm observation and diagnostics UX

- [x] Stop synthesizing duplicate observed values; keep legacy input and genuine observation-only state.
- [x] Replace snapshot-container support warnings with actionable captured warnings through shared validation.
- [x] Add a collapsed, accessible Designer problems panel with structured issues and source navigation.
- [x] Update ADR/specifications, user guidance and changelogs.
- [x] Run focused Go and extension checks and review compatibility and source preservation.

Validation: Go UserForm/LSP/CLI/file pull/push/pack suites, full VS Code suite,
Designer DOM tests, type checks, lint/docs/build and packaging pass. A copied
real workbook produces 213 rather than 344 YAML lines, preserves three actual
warnings, and leaves the original workbook unchanged. Live LSP preview assertions
pass for current/legacy YAML and Unicode JSON. Probe cleanup required terminating
its own LSP process after shutdown; native problem-navigation integration is not
directly tested. Evidence and commands:
`tmp_workspaces/userform-observed-20261017-e2e/verification.md`.

# PR #933 review follow-up

- [x] Validate checkbox/select submission against synchronous draft refs and regression tests.
- [x] Preserve a staged image when a replacement picker is cancelled or fails; cover host transactions.
- [x] Assess frm discovery errors against duplicate detection and selected-path link validation.
- [x] Run extension checks before pushing confirmed fixes to the existing PR.

Review disposition: checkbox/select changes synchronously update the draft ref
and existing DOM submission tests pass. frm discovery stays fail-closed because
an unreadable subtree can conceal a duplicate code target. Picture replacement
is fixed with host/Webview regressions covering cancel, failure, successful
replacement and late results after modal cancellation. Complete extension tests,
type checks, lint, docs and scoped formatting pass.

# Designer SpinButton appearance

- [x] Replace text arrows with two raised buttons filling the control.
- [x] Follow horizontal and vertical control proportions.
- [x] Verify visual appearance and rebuild the VSIX.

# Designer ScrollBar appearance

- [x] Replace the centered text glyph with end arrows, track and raised thumb.
- [x] Match horizontal and vertical auto-orientation to the preview dimensions.
- [x] Verify native layout, visual appearance and rebuild the VSIX.

# Designer Undo/Redo buttons

- [x] Inspect host availability and shipped toolbar icon rendering.
- [x] Verify actual mouse clicks against registered native Designer history.
- [x] Fix confirmed defects and update regression coverage and VSIX.

Validation: the bundled Codicon CSS lacks `undo`; the supported `discard` glyph
fixes the icon and the glyph regression fails before/passes after the fix. Native
VS Code tests exercise property text input, responsive pane navigation and actual
Undo/Redo mouse clicks. The full suite, type checks, lint, docs and packaging pass.
History availability follows confirmed Designer edits and native Undo/Redo state.

# Designer native double-click

- [x] Reproduce viewport-retargeted double-clicks missed by direct event dispatch.
- [x] Recover the visual target without changing pointer capture or drag behavior.
- [x] Verify native mouse double-clicks and rebuild the VSIX.

Validation: the focused regression failed before the fix and passed afterward.
The full VS Code suite passes with native Chromium control/form double-clicks
and drag assertions. Type checks, lint, docs, scoped formatting, and packaging
pass. Evidence: `C:\Users\HARUMI\orca\workspaces\xlflow\guppy\tmp_workspaces\userform-doubleclick-20261016\verification.md`.

# Designer event code navigation

- [x] Add an Excel-verified Go event catalogue and versioned LSP list/edit requests.
- [x] Connect Designer double-click and event selection to targeted, native-undoable sidecar/frm code edits.
- [x] Verify source publication, existing-handler navigation, stale requests and UI gestures.
- [x] Verify generated code with real Excel and update contracts, guides and changelogs.

Validation: full VS Code 1.142.0 suite, type checks, lint, documentation and
changed-file format checks pass. Native tests cover targeted saving, code Undo,
dirty buffers, both code modes, Page owners, unchanged preview refreshes and
source/connection invalidation. Go LSP/UserForm suites pass. Real Excel verified
all 222 declarations and representative runtime events, with saved worksheet
sentinels in sidecar and frm modes. Review findings about linked code paths and
existing-file creation were fixed; junction regressions pass. Excel sessions
stopped. Exact commands, evidence and limits:
`C:\Users\HARUMI\orca\workspaces\xlflow\guppy\tmp_workspaces\userform-events-20261016\verification.md`.

# Designer actions and sidebar form build

- [x] Add the source-editor context command with the existing FormSpec/resource boundary.
- [x] Reuse native Undo/Redo and existing subtree confirmation for toolbar history and Backspace.
- [x] Add individual/all form builds with targeted explicit saving, session reuse and partial-failure reporting.
- [x] Verify extension regressions and real Excel builds; update specifications, guides and changelogs.

Validation: full VS Code 1.142.0 tests, lint, documentation and changed-file
format checks pass. Real Excel 16.0 x64 passed individual/batch creation,
overwrite, existing-session reuse, partial-failure retention, nested snapshots
and code-behind runtime assertions. Review fixed build save targets to follow
the CLI's sidecar authority in both display modes. All test sessions stopped.
Exact commands, results and limits:
`C:\Users\HARUMI\orca\workspaces\xlflow\guppy\tmp_workspaces\designer-actions-build-20261016\verification.md`.

# Designer concrete property inputs

- [x] Remove scalar type/null selectors and typed-object null selectors while retaining metadata authority.
- [x] Preserve existing absent/null source states until concrete confirmation; keep false, zero and empty strings valid. For `valueType: any`, retain an authored concrete type or use text without type inference.
- [x] Remove only obsolete selector CSS and align specifications, Japanese/English guides and both changelogs; leave FormSpec/API null support unchanged.
- [x] Validate targeted property-grid regressions, extension checks and formatting before marking complete.

Validation: `test:designer` and full `pnpm run test` passed with VS Code 1.142.0,
including existing null scalar/object preservation and concrete-value edits.
Lint, docs, changed-file formatting and diff checks pass. Dark Japanese and
narrow light English displays were inspected; bounded read-only review found no
actionable issues. This UI-only change did not rerun Excel integration.

# UserForm typed property editors

- [x] Reset independent control Font presentation to prevent inherited family, size, Bold, Italic, Underline and StrikeThrough. Real Webview regressions cover Form/Frame, default and partial child Fonts, explicit child Fonts and source immutability; full VS Code 1.142.0 tests, lint, docs and formatting pass.

- [x] Prevent Form/Frame ForeColor inheriting into default child controls. Reproduced with computed-style assertions in real VS Code, then verified independent default/explicit/nested colors; full VS Code 1.142.0 tests, lint, docs and formatting checks pass.

- [x] Add VBE-style 48-color Palette/System choices, preserving modal-only previews, raw OLE input and metadata authority for system colors; verify selection, keyboard navigation, Cancel and atomic OK.

- [x] Add partial Font, semantic Color and enum contracts with backward-compatible metadata and source-preserving editing.
- [x] Implement shared binary read/generate/mutate/projection and preserve unedited state.
- [x] Add modal editor registry, live previews, native Image assets and atomic Position editing.
- [x] Verify old/new capability combinations, cancellation, stale drafts and native Undo/Redo.
- [x] Update ADR/specs/schema/guides/changelogs and validate session-backed Excel and pack artifacts.

Validation: full UserForm and LSP tests, focused CLI tests, vet, complete VS Code
1.142.0 tests, lint, documentation and formatting checks passed. Real Excel
16.0 x64 passed full appearance/image assertions, blank/template pack artifacts,
partial-update preservation and file push. Review fixed rollback reference checks
for unsaved buffers and closed Designer checkpoints. Evidence and exact commands:
`C:\Users\HARUMI\orca\workspaces\xlflow\guppy\tmp_workspaces\userform-property-editors-20261015-202945\verification.md`.
No remote CI or supervised final review was run.

# Designer form resize

- [x] Add selected form edge/corner drag handles and client-size preview with zoom, snapping and child bounds.
- [x] Support atomic width/height transactions through the Webview and LSP boundaries, preserving YAML/JSON source and Undo/Redo.
- [x] Verify cancellation, capability gates, native Webview hit testing and full extension checks; update specifications and changelogs. VS Code 1.141.1 full tests, LSP full tests (157.730s), vet, lint/docs and formatting pass. No Excel workbook E2E is needed for this source-only gesture change.

# Issue #922 Designer productivity

- [x] Confirmed Excel 16.0 twip/Single geometry rounding; limited tolerance to four geometry fields and added acceptance/rejection regressions (528 .NET tests pass). Rebuilt/installed both binaries. Original fractional FormSpec succeeds through build/push/run/save/pull/lint and saved-file apply; sentinel, 12 geometry values, unchanged spec and COM cleanup verified. Lint/format checks pass. Evidence: tmp_workspaces/final-review/issue922/coordinate-precision-evidence.md.
- [x] Reproduced interior click selection loss in native VS Code Webview before the fix. Restrict blank content to Frame; coordinate hit tests cover Label/TextBox, nested children, holding an existing selection and single/group movement. Full VS Code tests and lint passed.
- [x] PR #932: validate and fix six review findings (paste/container capability, Page Delete, live group boundary, context-menu dismissal/focus and tab-order parent choices); focused regressions, full VS Code tests, lint and diff checks passed.
- [x] Validate Pass 1 custom-control tab-order finding, align eligibility with Property Grid metadata and add YAML/JSON regressions.
- [x] Preserve Pass 1 evidence and clean up its review resources; FormSpec tests, focused LSP tests and vet pass.
- [x] Run independent Pass 2 on 568a86c: PASS; prior P2 resolved, no additional findings. Preserve report and clean up review resources.
- [x] Correct native Webview clipping-test viewport preparation after Linux CI failure; full VS Code tests, oxlint, formatting and diff check pass locally. Push correction to PR #932; remote confirmation remains pending.

- [x] Extend atomic source-preserving semantic operations and copy RPC for YAML/YML/JSON.
- [x] Add sibling selection, group layout, clipboard, z-order and Tab order UI.
- [x] Integrate capability negotiation, shared host clipboard and stale-safe edit queue.
- [x] Verify Go/RPC, Webview and native VS Code Undo/Redo regressions; review complete diff.
- [x] Record ADR/spec/user documentation and changelog updates.

Validation: `rtk pnpm --dir editors/vscode test` passed with VS Code 1.141.0,
including group gestures, shared clipboard, custom-editor Undo/Redo, reversed
Copy completions, stale source, capability gates, failed/cancelled Cut clipboard
retention and successful Cut publication. The full
`rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./internal/vba/userforms/spec/... ./internal/lspserver -count=1`
passed (LSP: 186.776s); final focused UserForm/productivity tests and matching
`vet` passed after backend refinements. `rtk task lint`, documentation checks,
docs build, changed-file formatting and `rtk git diff --check` passed.
Read-only host/UI reviews were completed; the confirmed custom-ProgID marquee
inconsistency was fixed and covered, and fractional group translation preserves
source precision. Excel/COM/VBE and remote CI were not run. No Excel
`tmp_workspaces` were used; this change edits FormSpec source only.

# Issue #921 final review follow-up

## PR #931 review comments

- [x] Restrict canvas/toolbar deletion to ordinary controls and cover Page-header/list/breadcrumb selection, while preserving Page-list deletion.
- [x] Remove duplicate Delete keys in both localization bundles and verify key uniqueness.
- [x] Describe the implemented Page/TabStrip feature in both changelogs and remove stale Issue #921 deferrals.
- [x] Verify Designer/full VS Code tests, docs/format/lint checks and localization key uniqueness; prepare the fixes for the existing PR.

- [x] Validate Pass 1 P2 against the accepted Designer capability contract and implementation.
- [x] Align MultiPage/TabStrip documentation: require geometry for both pointer and picker reparenting, explicitly name cumulative Page capability, and distinguish childful from empty Page deletion.
- [x] Verify documentation/format checks and preserve the independent review report and disposition.
- [x] Acknowledge the settled review worker; release retained its external terminal, so explicitly close review-only terminals and remove the worktree after preserving evidence.

# Issue #920 Frame parenting and hierarchy

## PR #929 review follow-up

- [x] Fix CHANGELOG continuation placement after validating the Markdown grouping.
- [x] Carry preview ProgID and share the built-in Frame predicate across targeting, placement and host preflight; reject custom destinations before dispatch.
- [x] Cover implicit/matching/custom ProgIDs and custom-parent host rejection; full extension suite and `task lint` PASS. Push to the existing PR after commit checks.

- [x] Validate Pass 1 P2: visible Frame overflow disagreed with drop hit testing.
- [x] Align container rendering with content-bound hit tests and preserve source coordinates; add real Webview clipping and short-drag regressions.
- [x] Verify review fix with `rtk pnpm --dir editors/vscode test` and `rtk task lint`; both PASS.
- [x] Preserve Pass 1 report/probes in `tmp_workspaces/final-review/issue920/` and remove its review worktree and four review-specific terminal tabs.
- [x] Run Pass 2 against `2fb768e`: PASS, no confirmed findings; reviewer independently reran the full VS Code suite. Preserve report and remove Pass 2 review resources.

- [x] Extend source-preserving YAML/JSON parenting and enforce Designer RPC scope.
- [x] Negotiate parent-edit capability and preserve stale-document/Undo guards.
- [x] Add container targeting, world-space drag previews, Frame placement and breadcrumbs.
- [x] Verify engine, host, DOM and real VS Code regressions; review the complete diff.
- [x] Record parenting decisions/specifications and update public docs/changelogs.
- [x] Review follow-up: return focus to the canvas after breadcrumb selection and verify keyboard movement in DOM and real Webview tests.

Validation: `rtk pnpm --dir editors/vscode test` passed with real VS Code
1.141.0 Webview parenting/placement, coordinate alignment, breadcrumb keyboard
selection and native Undo/Redo. The complete UserForms/LSP command
`rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./internal/vba/userforms/... ./internal/lspserver`
passed; the uncached final LSP run took 129.871s. Matching `vet`, `rtk task lint`,
changed-file Oxfmt/Go formatting and `rtk git diff --check` passed. The focused
read-only UI review found the breadcrumb focus issue, now fixed and covered.
Excel/COM/VBE and remote CI were not run; this feature edits FormSpec source
and does not change workbook-backed behavior. No Excel `tmp_workspaces` were used.

# PR #928 review follow-up

- [x] Reproduce Delete bubbling from a focused Page tab and custom ProgID insertion through the Designer RPC.
- [x] Restrict canvas Delete to the canvas target; constrain explicit ProgID to the advertised built-in contract without changing shared semantic edits.
- [x] Verify regressions, existing UI/Undo and Go contracts; prepare the fixes for the existing PR.

# Issue #919 final review follow-up

- [x] Reproduce Pass 1 RPC bypass of root-only insertion and Page exclusion.
- [x] Enforce the Designer structural scope at the LSP boundary while preserving the shared edit API.
- [x] Verify YAML/JSON rejection, subtree deletion and existing edit contracts; prepare the verified fixes for Pass 2.
- [x] Preserve the Pass 1 report and complete review resource cleanup.

Pass 1 P2 verified and fixed at the Designer RPC boundary. Focused UserForm edit
tests passed; full related UserForm/LSP tests passed (LSP 127.468s), along with
vet, lint, docs and formatting. Generic parent/Page/custom edit API and allowed
MultiPage cascade/child deletion are protected by regression tests. Report and
cleanup evidence: `tmp_workspaces/final-review/issue919/pass-1-disposition.md`.

# Issue #919 Toolbox and structural editing

- [x] Add shared control metadata, click placement and stable ID/name generation.
- [x] Expose source-preserving YAML/JSON add/delete via structural LSP capability.
- [x] Confirm subtree deletion in the URI edit queue and retain stale/undo guarantees.
- [x] Verify generation, UI, compiler, RPC, native undo/redo and source preservation.
- [x] Document the contract, update changelogs and complete self-review.

Validation: `rtk pnpm --dir editors/vscode test` passed, including real VS Code
1.140.0 Webview and native Undo/Redo. Related Go checks passed with
`rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./internal/vba/userforms/... ./internal/lspserver`
(LSP: 132.554s); the corresponding `vet` also passed. Lint, docs checks,
changed-file Oxfmt, Go formatting and `git diff --check` passed. Focused read-only
UI and Go reviews found no confirmed defects. Excel device tests and remote CI
were not run; this change edits source documents and does not alter COM/VBIDE.

# PR #927 review follow-up

- [x] Reproduce and isolate form/control draft identity when a control ID is `form`.
- [x] Restore canonical common scalar edits/metadata for custom ProgID controls without exposing type-specific fields.
- [x] Reject property number payloads whose scalar conversion would change their decimal value; retain exact int64 and ordinary decimal edits.
- [x] Verify source/RPC/DOM regressions, full related Go/LSP (134.548s), real VS Code 1.140.0, lint/docs, vet and formatting; prepare the verified fixes and evidence replies for the existing PR.

# Issue #918 final review follow-up

- [x] Pass 2: validate independent property capability at the client request boundary and protect authored numeric precision at the Webview boundary.
- [x] Verify both regressions, preserve Pass 2 report and cleanup resources; stop the full review loop at two passes. Full spec/LSP, real VS Code 1.140.0, lint/docs, format and vet pass; the final corrections receive coordinator verification, not a third independent review.

- [x] Reproduce hidden-panel property reply loss; retain/replay edit results without applying hidden stale edits.
- [x] Verify source safety, no-op/error/success replies, resumed Property Grid and real VS Code regressions.
- [x] Preserve Pass 1 evidence, clean up review resources and prepare the verified fix for bounded Pass 2.

# Issue #918 Property Grid

- [x] Extend YAML/JSON scalar editing and expose source-based property metadata through LSP.
- [x] Add localized Property Grid with shared selection, draft validation and semantic commits.
- [x] Verify source preservation, real Webview behavior, native undo/redo and geometry regressions.
- [x] Record the design contract, update user documentation/changelogs and complete self-review.
- [x] Review follow-up: preserve empty multiline JSON build layout and retain rejected drafts/field diagnostics across selection changes; verify regressions and the full VS Code suite.

# PR #926 review follow-up

- [x] CodeRabbit: reproduce secondary pointerdown replacing the active gesture; reject new pointerdown until that gesture ends and verify primary commit.
- [x] Verify full VSCode/lint/docs/format checks and publish the correction with an evidence reply.

- [x] Verify empty Frame/MultiPage minimum size and secondary-pointer cancellation with regressions.
- [x] Apply child extents plus insets separately from the outer 1pt minimum; preserve active pointer ownership for cancellation/capture loss.
- [x] Run VSCode/lint/docs/format checks; publish the verified fixes on existing PR with evidence replies and persisted body verification.

# Issue #917 final review follow-up

- [x] Preserve unchanged axes during move/snap and reject gestures that leave existing overflow; add geometry regressions.
- [x] Enforce a single move/resize operation or distinct move+resize pair for one control at the RPC boundary; test rejected raw RPC batches and accepted pairs.
- [x] Run affected Go and VS Code checks, preserve Pass 1 evidence and clean up review resources.
- [x] Commit the fixes and run independent Pass 2 against the full Issue #917 change; both Pass 1 findings resolved.
- [x] Reject moves of existing sub-point controls; verify explicit resize repair and real Webview no-edit behavior.
- [x] Preserve Pass 2 report and clean up its resources; stop the full review loop at two passes and verify the isolated P2 correction through regression tests/self-review.

# PR #925 review follow-up

- [x] 操作エラーの中間状態検証による上書きを防ぐ。
- [x] ID 重複診断を canonical validator に委譲し、追加対象 ID の文脈を保持する。
- [x] 複数行引用 scalar のコメント二重化と indentless sequence の最後の要素削除を修正する。
- [x] reorder は整数 gap を利用し、必要な最小連続範囲だけを変更する。
- [x] 回帰テスト・vet・docs チェックを確認する。

# Issue #916 final review follow-up

- [x] 明示タグ付き引用 scalar の境界検出を修正する。
- [x] 引用内の #、エスケープ、flow 区切り、LF/CRLF のコメント保持を回帰テストする。
- [x] 関連 Go テストと vet を確認し、Pass 1 の報告と再現コードを保持してレビュー環境を cleanup する。
- [ ] 修正をコミットして独立レビュー Pass 2 を実行する。

# Issue #915 designer review follow-up

- [x] PR #924: prefer client/inside dimensions on both axes; render explicit ComboBox values including empty strings; localize Webview UI through host l10n messages.
- [x] Preserve zIndex-ordered MultiPage selection to match pure-Go generation/topology; add reordered Page regression and explain unsupported source-order recommendation.
- [x] Run renderer and real Webview tests, lint/docs checks; publish with regression evidence. Docstring coverage is advisory, not a repository gate.

- [x] Honor explicit visibility for controls and Page tabs, retaining original Page collection indexes.
- [x] Add hidden-parent and hidden-Page selection regressions and rerun Designer/Windows Webview checks.

# PR #923 review follow-up

- [x] Resolve picture roles before every canonical VBA component collector; cover module, class and document roots through CLI pack/file push.
- [x] Preserve case-distinct Code/ loose modules and fingerprints on case-sensitive hosts.
- [x] Correct Windows CI assertions to resolve expected physical paths before host-key comparison, retaining symlink target checks. Directory-alias fixtures cover resolved roots as well as filename casing; the initial case-only correction was insufficient on the runner.
- [x] Remove the CodeQL allocation-size sum, propagate input path-resolution errors, and correct obsolete pack restrictions.
- [x] Run affected Windows/Linux checks, lint and installed Excel gate; publish fixes and report review dispositions. CodeQL alert 11 is fixed and Linux CI passes; Windows CI is rerun after the physical-path assertion correction.

# PR #913 review follow-up

The persisted-caption finding is valid: caption projection respects the
control contract, but unsupported-property reporting still treats every
Caption as projected and hides nonempty unmodeled values.

- [x] Add full-projection regressions for nonempty unsupported captions,
      empty defaults, and supported captions; confirm the regression fails.
- [x] Align unsupported Caption reporting with the canonical control contract
      while keeping empty unsupported defaults silent.
- [x] Verify affected Go tests, lint/docs and staged formatting.

Publish the verified fix to the existing PR and reply with the regression
evidence. Test-function docstrings are an advisory suggestion rather than a
repository requirement and are outside this behavior fix.

# PR #907 review follow-up

All five comments are valid presentation/projection defects; retain global
resolution, native metrics identities and incident-edge selection.

1. [x] Filter displayed root candidates while retaining unresolved requests.
2. [x] Retain all related SCC member dependency nodes and boundary IDs.
3. [x] Join conditional procedure metrics by declaration start position.
4. [x] Allow existing empty directories as display scopes; reject invalid files.
5. [x] Distinguish the field inventory from mutable access counts in human output.
6. [x] Add regressions and verify architecture/output/CLI tests, affected race,
       lint/docs/staged formatting and whitespace checks.

# Issue #463 final-review Pass 1 follow-up

Validated against `d143e7e8`: implicit Application callbacks are extracted
from unresolved IR and can lose target evidence and possible reachability.

1. [x] Project callback references from the existing resolved IR snapshot.
2. [x] Cover receiverless and With Application Run/OnTime/OnKey, including
       project-procedure and non-callable local shadow controls.
3. [x] Run focused architecture/calls/reachability/CLI tests, race and lint.
4. [x] Pass 2 reviewed `4a0fda3c`: the P2 is resolved, the focused callback
       regression passes independently, and no confirmed findings remain.

# PR #906 review follow-up

CodeQL #10 identifies unchecked addition in a capacity hint; use the existing
raw length directly. The ADR's graduation boundary is historical and must be
identified as superseded by the later UserForm amendments. The repeated-root-
Caption report is not reproduced: applyPersistenceEdit updates the clone's
Caption before the next iteration, confirmed by the exact reported sequence.

1. [x] Remove the allocation-size sum and verify caption byte preservation.
2. [x] Cover repeated root Caption edits and classify the review with evidence:
       Original -> Changed -> Original passes without changing ApplyEdits.
3. [x] Mark historical ADR authority consistently with current amendments.
4. [x] Run full/focused tests, affected race/vet, docs and formatting checks.

Publish the verified changes to PR #906 and reply with each classification.

# Issue #887 final-review Pass 1 follow-up

Validated against `c87cac40`: the cleanup PID race, workspace reparse-point
boundary and unstable inventory errors are valid and in scope. Duplicate
MSForms-reference handling is unchanged from the parent commit; its VBE
acceptance remains unverified and is a separate follow-up, not a review fix.

1. [x] Pin the verified process handle through cleanup; never kill by PID.
2. [x] Reject existing reparse-point ancestors before creating a workspace.
3. [x] Sort inventory validation keys and add multi-error regressions.
4. [x] Run focused/full verification and the real Excel gate.
5. [x] Commit fixes and complete the bounded Pass 2 review.

Pass 2 reviewed `8af2ff79` with gpt-6-luna xhigh: PASS, all three in-scope
findings resolved, no new confirmed findings. Independent harness and
inventory regressions passed; retained Excel evidence was read back.

# Issue #887

- [x] Add explicit template/source UserForm topology configuration (default template).
- [x] Discover canonical template specs alongside legacy code-only forms; retain explicit property intent.
- [x] Plan supported Designer updates/additions/removals atomically with lossless preservation.
- [x] Reconcile project metadata and reparse output before publication.
- [x] Cover Go/CLI regressions and run focused/full/race/vet/lint checks.
- [x] Verify template and blank artifacts in real Excel, including save/reopen and final-form removal.
- [x] Update ADR/spec/config docs/CHANGELOG and review the complete diff.

Local commands, retained workspaces and results:
`internal/pack/testdata/userform-integration/README.md`.

# PR #904 review follow-up

REGISTERED aggregate-size and malformed Forms LIBID findings are valid.
PROJECT has an aggregate Size (the comment's first-length explanation is
incorrect), but it too must be ignored on read in favor of two bounded paths.

1. [x] Share bounded reference payload sizing between both parser paths.
2. [x] Validate LIBID grammar and cover malformed Forms admission atomically.
3. [x] Cover PROJECT and REGISTERED aggregate-size variations, truncations
       and reference byte preservation; remove the unused E2E workbook table.
4. [x] Affected pack/sourceinventory/CLI tests, race, vet, parser fuzz,
       lint/docs/PowerShell and staged formatting checks pass.

Publish the verified fix to the existing PR and reply with these results.

# Issue #886 final-review Pass 1 follow-up

Validated against `852a57b7`: the P2 NUL validation finding is in scope.

1. [x] Validate every decoded reference string, including ORIGINAL, CONTROL
       twiddled LIBID and extended name, before accepting a reference group.
2. [x] Add corrupted-field regressions and atomic Forms admission coverage;
       retain valid Excel reference byte-preservation controls.
3. [x] Run pack tests, race and vet; commit hooks check lint and formatting.
4. [x] Pass 2 reviewed `eb21a91f` with gpt-6-luna xhigh: the P2 is resolved,
       no confirmed in-scope defects remain, and affected package tests pass.

# Issue #883 final-review Pass 1 follow-up

Validated against `76cae6b8`: both P2 findings are in scope.

1. [x] Require canonical MSForms LIBID evidence in raw reference records;
       reject an unrelated twiddled GUID even if its display name is MSForms.
2. [x] Reject equal VB_Base GUIDs at the project-addition boundary, including
       case variants, without mutating the input project.
3. [x] Add regressions and run focused userforms/pack/excel tests, vet and docs
       checks; commit hooks run lint and staged formatting checks.
4. [x] Obtain Pass 2 with gpt-6-luna xhigh: `ace0f284` passed with both P2 fixes
       verified and no new findings. Preserve #886/#887 scope. The admission
       checks do not change generated Designer bytes;
       committed Excel normalization evidence still passes in ordinary tests.

# PR #902 review follow-up (Issue #882)

Verified against `40267db2`: both Devin findings are valid. The two CodeQL
alerts refer to the same unchecked sum used for key-slice capacity.

1. [x] Reject invalid before/after coordinate systems before interpreting
       geometry; retain omitted, points and parent-relative inputs.
2. [x] Check the supplied legacy caption baseline when build.caption first
       appears, retaining changed-build precedence and atomic failure.
3. [x] Collect the property-key union without adding untrusted map lengths;
       cover overlap, disjoint keys and deterministic edits.
4. [x] Add focused regressions, update compiler/trusted-Excel contracts, and
       run focused/full tests, race/lint/docs and the saved/reopened Excel gate.

Publish the verified fix to the existing PR and reply to each review thread
with its disposition and evidence; keep remote CI status separate.

CodeRabbit's pack integration warning belongs to explicit follow-up #887;
private-helper docstring coverage is not a repository gate. The developer
Excel harness intentionally executes a sentinel, so explicitly document that
only trusted generated workbooks and known derivatives may be verified.

# Issue #882 final-review follow-up

Independent review of `d63cde3a` confirmed one in-scope P3: seven new
snake-case test labels are incorrectly extracted as structured error codes.
Keep this fix local to the new test labels rather than changing the existing
inventory generator's repository-wide extraction contract.

1. [x] Rename the seven test labels and regenerate the reference inventory.
2. [x] Run focused userforms/pack tests and docs/format/lint checks; confirm
       only the four real `userform_edit_*` codes remain newly inventoried.

Publication follows verification: commit the review fix, push the feature
branch, create the PR, and validate its GitHub read-back.

# PR #901 review follow-up (Issue #881)

Verified against commit `7396d1fe`; all three Devin bug reports are valid and
in scope. The temporary-allocation observation is a performance hardening note,
not a demonstrated regression, and remains a separate follow-up.

1. [x] Resolve known Designer stream names with CFB case-insensitive identity,
       retain their original spelling, and replay every captured stream under
       that spelling. Add case-varied root and optional-stream round-trip tests.
2. [x] Let `pull --backend auto` inspect module topology without parsing
       Designer data when `code_source = "frm"`, so unsupported Designer
       layouts can still select Excel; keep strict parsing for the file backend
       and sidecar probe. Add malformed-Designer probe regressions.
3. [x] Count parsed Designer streams in `pack.modules.carried_streams` and add
       a form-bearing output-contract regression.
4. [x] Run focused and full validation, commit and push the fixes, then reply
       to all review threads with evidence and the allocation disposition.

# PR #900 review follow-up

Verified against commit `c4ed68d8`; the stale compatibility artifact portion
of the Devin finding is valid and in scope. Direct FormSpec-to-MS-OFORMS push
remains assigned to parent Issue #876.

1. [x] Mark specs emitted by `pull --backend file` as having unsynchronized
       compatibility `.frm` / `.frx` artifacts.
2. [x] Make the shared Excel/file push preflight reject that persistent marker
       even when a same-named stale `.frm` exists; keep the artifacts
       non-destructive and unmanaged.
3. [x] Add focused fresh/stale artifact regressions, update ADR/spec/docs and
       the real-Excel gate, and run validation. The review reply must preserve
       the Issue #876 scope boundary.

# PR #893 review follow-up 2

Verified against commit `4c57c077`; the stale delegated warning finding is
valid and in scope:

1. [x] Explicitly clear `XLFLOW_PULL_AUTO_PROBE_WARNING` for warning-free WSL
       auto-pull delegation, even when the parent environment and `WSLENV`
       contain a stale value.
2. [x] Add focused delegation regression coverage, run repository validation,
       push the fix, and reply to the review thread with evidence.

# PR #893 review follow-up

Verified against commit `bb327504`; three inline findings are valid and in
scope, while WSL detection of arbitrary externally opened Excel workbooks is
the approved local-auto boundary rather than a defect:

1. [x] Skip workbook coordination before `auto` selection, then acquire the
       workbook lease only when auto actually selects Excel. Keep file pulls
       under source-tree coordination and preserve explicit backend behavior.
2. [x] On an indeterminate Windows open-state probe, require attachment to an
       already-open matching workbook instead of opening a separate saved copy.
3. [x] Preserve the WSL session-probe failure warning in the delegated Windows
       JSON result, with focused end-to-end delegation coverage.
4. [x] Investigate the VS Code CI failure independently, run focused/full
       validation, push the review fix, and reply to every inline thread with
       evidence or the documented design disposition.

# Issue #892 final-review pass 1 follow-up

Independent review of commit `4e08cf77` confirmed one in-scope P2:

1. [x] Restrict the auto backend capability probe to saved-workbook
       properties (parse/protection/UserForm support) so it never reads managed
       source trees before source-tree leases and `--wait` take effect.
2. [x] Keep source-path, stale-file, and publication planning validation inside
       `PullContext` while its leases are held; add focused regression coverage.
3. [x] Run focused/full tests and lint, commit the fix, and request final-review
       pass 2 on the exact fix commit.

# Issue #877 final-review pass 1 follow-up

Independent review of commit `d20ca701` confirmed one in-scope P1:

1. [x] Key every CFB sibling tree by `DirectoryNameKey` so case-variant
       storage definitions deduplicate only when metadata matches, conflicting
       metadata and storage/stream collisions fail deterministically, and the
       first spelling remains the serialized display name.
2. [x] Reject case-insensitive storage/storage and storage/stream duplicates
       while reading malformed CFB directory trees.
3. [x] Add writer, reader, determinism, and malformed-input regressions; run
       focused/full tests, lint, format checks, and the real-Excel pack gate.

# PR #874 review follow-up (Issue #871)

Verified against commit `f5bef996`; the five inline findings and one
documentation finding are valid and in scope. The generic docstring coverage
warning does not identify a repository contract violation and needs no change.

1. [x] Preserve case-only module rename semantics on case-sensitive filesystems
       without deleting an aliased target on case-insensitive filesystems.
2. [x] Classify wrapped publication failures as `pull_source_publish_failed`
       before generic workbook-path errors.
3. [x] Reject overlap in either direction between the forms root and each
       managed source root before reconciliation.
4. [x] Keep `target.kind` within the stable `file` vocabulary and make pull
       prerequisites backend-specific.
5. [x] Add focused regressions, run affected tests and lint/docs checks, commit,
       push, and reply to each review thread with evidence.

# PR #864 review follow-up (Issue #855)

Verified against commit `b8d8a2e5`; both Devin comments are valid and in
scope:

1. Preserve the existing `pack` contract that configured absolute source
   roots may live outside the project while keeping the containment contract
   of ordinary encoding commands unchanged. Add a pack-only sourceencoding
   option, retain managed-root symlink containment, and report external
   diagnostic paths as absolute paths.
2. Normalize configured source-root separators before discovery so Unix
   preflight and `sourceinventory` consume the same paths instead of silently
   skipping backslash-configured roots.
3. Add sourceencoding and pack CLI regressions for valid/invalid external
   roots and invalid source under a backslash-configured root. Run focused and
   full tests, lint/docs/format checks, push the follow-up commit, and reply to
   both review threads.

# PR #863 review follow-up (Issue #854)

Verified against commit `0e6c6145`; the six Devin observations and two
CodeRabbit comments reduce to six valid in-scope fixes:

1. Publish through the canonical output path so an existing output symlink is
   preserved and its referent is replaced; retain both lexical and canonical
   candidates for Office lock-file detection. Extend identity comparison with
   `os.SameFile` for existing hard-link aliases.
2. Create staging files exclusively with mode `0644` so Unix umask applies to
   new outputs, and preserve an existing destination's permission bits during
   replacement.
3. Keep the exclusive replacement probe on Windows, where sharing modes are
   meaningful, but do not require write access to the old file on Unix; final
   rename errors remain the source of truth there.
4. Use Linux `renameat2(RENAME_NOREPLACE)` for atomic create and retain hard
   links only as the ENOSYS/EINVAL fallback and for other non-Windows systems.
   Never classify create-time EPERM/EACCES as a busy destination.
5. Surface cleanup failure together with pre-publication errors while keeping
   `errors.Is` classification intact.
6. Read the staged `xl/vbaProject.bin` entry fully so ZIP CRC/data corruption
   is detected before publication.

Add focused Windows/Unix/CLI regressions, run affected tests on Windows and
WSL Linux, run lint/docs checks, commit, push to PR #863, and reply to each
review thread with the disposition.

# Final review pass 3 follow-ups (validate-static-rules-default, aa889dfc)

Independent whole-branch review verdict: approve-with-notes. Remaining items:

1. DONE (PR #860 review round 4): review_test.go now bounds Unreviewed by
   the committed backlog (<= 9231). Review-driven decreases still pass;
   new snapshot findings fail until triaged or the ceiling is raised.
2. Follow-up: addressOfEntryNames is a name-only project-wide set; a
   same-named non-callback procedure can be over-suppressed (bounded
   false negative, consistent with the suppression-only fail-open
   contract). Revisit only if a real misreport is observed.

# Final review pass 1 fixes (validate-static-rules-default, cc016ad9)

Independent review verdict: approve-with-notes. Verified findings:

1. P2 cross-module AddressOf signature constraint: VBA resolves a bare
   AddressOf target project-wide across public procedures in standard
   modules, but parameterPassingSignatureConstrained only consulted the
   candidate file's own facts.addressOfNames. A callback declared in a
   dedicated module still received VBA273/VBA275/VBA277 advice - the same FP
   class this branch fixed in-module. VBA265 is unaffected (its candidates
   are Private-only) and keeps its current merged dynamic-entry contract.
   Fix: add analysisContext.addressOfEntryNames (project-wide union of
   facts.addressOfNames) populated when any signature-constrained
   parameter-passing rule is enabled (VBA273/275/276/277), thread it into
   parameterPassingSignatureConstrained alongside the file-local check, and
   widen it in the realtime path from projectDocuments the same way
   dynamicEntryNames is widened. Add a two-module regression test.
2. P3 metrics assertion: review_test.go now uses floors + dropped
   Unreviewed==0. Intentional (opt-in profile emits unreviewed observations
   by design); add a TODO documenting when the bounds can be tightened.
3. P3 missing config row: vitepress/reference/config-file.md [analyze]
   table lacks detect_invalid_ismissing_usage (pre-existing gap, now
   user-facing because VBA283 is default-on). Add the row + narrative line.
4. P3 stale comments/doc: analyzer.go dynamicEntryNames comment (mentions
   VBA260, actual consumers are VBA265 + now includes AddressOf names),
   module_facts.go "three source scans" comment (now four), and
   vba-parameter-passing-diagnostics.md VBA276 exclusion list (code applies
   the AddressOf constraint too, though unreachable for Property Let/Set -
   document the check or the dead branch).
5. Unsupported/follow-up observations recorded by the reviewer (realtime
   private-name-collision parity, Application.Run string-literal names not
   consulted by parameter passing) are pre-existing design boundaries; the
   realtime addressOfEntryNames widening shares the VBA265 scan loop so
   parity bookkeeping stays symmetric. No separate fix.

# Issue #826: VBA class/interface public-API hazard diagnostics (VBA278-VBA282)

Parent issue #816. Five opt-in warning rules, all default-disabled,
warning-level, inline-suppressible, non-blocking, on batch + realtime + LSP.
Renumbered from the originally implemented VBA273-VBA277 because origin/main
PR #845 (issue #823) already owns those IDs for parameter-passing rules.

- VBA278 public member underscore names (class/form/document; skips
  control-verified/intrinsic event handlers, verified Implements bindings,
  unresolved-Implements prefixes, and Class_Initialize/Class_Terminate).
- VBA279 non-Private Enum in document modules.
- VBA280 write-only Property (Let/Set without Get), one finding per name;
  Friend counts as public-facing; conditional/recovered accessors mark the
  name uncertain and fail open.
- VBA281 public members whose <Interface>\_<Member> binding resolves to a
  declared public member of the implemented interface class, and explicitly
  Public recognized event handlers.
- VBA282 self-name access inside predeclared modules (document/form/
  VB_PredeclaredId class); local/parameter/module-scope shadows and TypeOf
  type operands excluded; recovered/conditional procedures skipped.

PR #847 review fixes applied (Devin 1-6 + CodeRabbit 1-4; all verified valid):

- Form helper misclassified as event: recognizedEventMember validates the
  control prefix against designer controls when known and falls back to
  intrinsic UserForm\_\* names when unknown. Lazily resolves controls because
  VBA220's gate only covers \_change/\_click names.
- IFace_Utility FP: interfaceBindingMatch requires the suffix to name a
  public member of the resolved, unambiguous interface class; resolved but
  absent suffixes fall back to VBA278, unresolved interfaces fail open.
- VBA280 uncertain accessors: Recovered/ConditionalBranches mark the
  property name uncertain; Friend writers now count as public-facing.
- VBA282 guards: proc.IR.Symbol.Recovered/ConditionalBranches fail open;
  ScopeLocal/ScopeParameter/ScopeModule accesses excluded; TypeOf exclusion
  narrowed to the last (type-operand) child of type_of_expression.
- Spec fail-open text updated to match implementation.

# Conflict resolution (merge origin/main @ aabb6575)

- Rule ID collision: our VBA273-277 -> VBA278-282; config keys unchanged.
- registry.json: main entries kept, ours appended renumbered.
- diagnostics.md regenerated from registry; oracle fixture dirs + ids and
  fixture_contract_test.go renumbered; workspace_test.go keys renumbered.
- Corpus snapshots to be regenerated (corpus:update-snapshots) and ledger
  rows renumbered to VBA278-282; metrics numbers rechecked after regen.

Remaining: corpus regen + metrics, format:check/docs:check, full test pass,
commit merge, push, update PR #847 body (new rule IDs + oracle case IDs
vba278-public-underscore-member, vba279-document-public-enum,
vba280-write-only-property, vba281-public-event-handler,
vba281-public-interface-member-collision, vba282-predeclared-self-access).

# PR #845 review follow-up (Issue #823, VBA273-VBA277)

Review comments on internal/analyze/parameter_passing.go. Verified against
code; both bug reports are valid and in scope. Plan:

1. [Devin + CodeRabbit] Property accessor collision: mutation records are
   keyed by qualified name only, so Property Get Item and Property Let Item
   share one summary (first registered wins). VBA275 can falsely claim the
   Let index parameter is never written; VBA274 can miss a written Let value
   parameter. Fix: kind-qualify record keys (kind|name) for summary storage
   and finding lookup; call-edge resolution uses candidate kind when present
   and falls back to unique-name match or possiblyWritten when ambiguous.
   Regression test: Property Get before Property Let, Let writes index and
   assigns value parameter.
2. [Devin] applyArgumentMutation checks parameterTypeIsObject before
   parameterTypeIsUDT, so a member argument on a UDT that shadows a builtin
   object name (Type Collection) drops the propagated write and VBA275
   falsely fires. Same latent issue family as the recordParameterMemberWrite
   fix. Fix: UDT check first, matching VBA name-resolution precedence.
   Regression test: module UDT named Collection, member forwarded to a
   writing ByRef call.

Advisory (Devin): no executed VBE oracle case IDs. Oracle v1 supports
compile probes only; the modeled semantics (property value ByVal, forced
ByVal parens) are covered by byref-parenthesized-variable (accepted) and
property-signature-valid/invalid at compile level; runtime cases are out of
scope for the current probe contract. Rules remain opt-in and fail open.

# Issue #879 final-review follow-up

- [x] Keep TabStrip `ListIndex` out of unsupported FormSpec fields and report it as unsupported; cover Project -> WriteSnapshot -> LoadFormSpec.
- [x] Report non-empty ControlSource and RowSource site state as unsupported, and audit all decoded site strings so meaningful unprojected state cannot disappear silently; add focused warning regressions.
- [x] Run focused userform projection/spec tests and the relevant package validation; review the complete diff.

# PR #897 review follow-up

- [x] Report non-default unsupported nested `Level` state on its owning container, and report unmodeled non-default site/control bitfield state without flagging file-format defaults; cover root and nested fixtures.
- [x] Surface non-default form `BooleanProperties` through the existing unsupported warning contract and add a regression.
- [x] Correct the YAML parentId completion test cursor offset; confirm Page/MultiPage remain projection/snapshot types rather than additions to the built-in authoring contract, and production CLI wiring belongs to parent issue #876.
- [x] Add and run same-workbook Excel snapshot versus pure-Go projection parity verification; run focused and affected package tests. Commit/push fixes and reply to review threads with evidence.

# PR #837 review follow-up (Issue #822, VBA257/VBA258)

Review comments on internal/analyze/discarded_return.go. Verified against IR
dump and resolver source; all six are valid. Plan:

1. [Devin] Parse errors permit false all-discard claims: VBA258 claims every
   caller discards, but a parse error can hide a call or a dynamic-dispatch
   reference (Application.Run reaches Private procedures too). Fix: bail out
   of the whole pass when any file.IR.Parse has HasError/HasMissing.
2. [CodeRabbit] PathFilter removes modules before parsedFiles; hidden modules
   can contain callers or dynamic references. Note: resolutionComplete is too
   broad (AnalyzeProject always sets typeDBResolutionIncomplete), so gate on
   PathFilter==nil instead. Combined with #1 this also covers the Friend case.
3. [Devin] Same-name arguments hidden by isCalleePart: 'Foo Foo' argument is
   skipped because it shares the callee name inside the call range. Fix:
   locate the exact callee token range (bytes.Index of Callee.Text inside
   call.Range) and exempt only expressions inside that range.
4. [CodeRabbit] Implicit Application dispatch: receiverless 'Run "x"' and
   With-block '.Run' produce no dynamic references. Fix: in calls package,
   map callee member run/ontime/onkey with nil/empty/"." receiver to
   "application.<member>" when resolution ran and did not bind a project
   procedure (keeps inspect path unchanged via ResolutionNotAttempted).
5. [Devin] Underscore exclusion too broad: Implements modules exclude every
   underscored proc. Fix: collect implemented interface names from
   TypeReferences (kind=="implements") and exclude only <Interface>\_<Member>.
6. [CodeRabbit nitpick] Implements exclusion test cannot fail (unknownDynamic
   suppresses everything and IWorker_DoWork is never called). Fix: add an
   independent test with a standalone IWorker_DoWork call and no dynamic
   dispatch.

Non-actionable: CodeRabbit docstring-coverage warning (bot threshold, repo
convention does not doc-comment every function).

# PR #837 review follow-up 2 (post-merge)

Second-round comments verified against merged HEAD (bfa7e23a). Devin's three
comments and CodeRabbit's two actionable comments are stale: they target the
pre-fix head and are already covered by 92cfacb3 (parse-error bail-out,
calleeTokenRange, implementedInterfaceNames, projectViewComplete gate,
implicitApplicationAPIName) - CodeRabbit marks both as addressed.

One valid finding remains:

1. [CodeRabbit] Interface names containing underscores: `Implements I_Foo`
   makes the implementation `I_Foo_Bar`, but strings.Cut splits at the first
   underscore so qualifier `i` never matches `i_foo`. Fix: prefix-match the
   lowered symbol name against each complete interface name + `_`. Add a
   regression test with an underscored interface name.

# PR #838 review follow-up (Issue #827, VBA260/VBA261/VBA262)

Verified against commit `2ebcf3f6`; actionable plan:

1. [Devin] Exclude bracket-escaped identifiers that bind to procedure, module,
   or project declarations from VBA262. Add declared enum/local/module and
   undeclared host-expression regressions.
2. [Devin] Require VBA260's `ThisWorkbook` token to be an unqualified,
   unshadowed project-global receiver. Reject member-qualified and locally
   shadowed forms with focused tests.
3. [Devin/CodeRabbit] Preserve logical-to-physical source positions for VBA260
   continuation statements so ranges and line suppressions bind to the actual
   physical line. Cover indentation and a match beginning on a continued line.
4. [CodeRabbit] Distinguish partial worksheet catalogs from unavailable
   catalogs in structured warning text while retaining valid mappings.
5. [CodeRabbit] Skip `sheetData` subtrees through `xml.Decoder.Skip` to reduce
   token handling while still validating XML through EOF and rejecting later
   conflicting `sheetPr` metadata.
6. [CodeRabbit acceptance check] Extend VBA261 diagnostics and tests to explain
   the error-behavior difference as well as binding and return-type semantics.
7. Run focused tests, full affected packages, corpus snapshots, lint/docs, and
   push the review-fix commit. Re-read PR comments and remote checks.

Non-actionable/follow-up: repository style does not require doc comments on all
private helpers. The security job reports GO-2026-6452 in the pre-existing
excelize v2.11.0 dependency with no fixed version, so it is tracked separately
from this review fix. VBE-oracle need will be reassessed after checking existing
bracket fixtures; VBA262 is not compile-equivalent and will not be promoted as
compile evidence.

# Issue #823 final-review pass 1 follow-up

Review run_7c453f1dc6a5 on commit 202af7ff. All four confirmed findings
verified against IR probes and real Excel runtime checks; all valid. Plan:

1. F1: With-block implicit member writes invisible to the mutation summary.
   Fix: build a shared withReceiver map (statement ID to innermost With
   receiver expression), resolve implicit member targets and implicit member
   call arguments through it, and fail open on the unresolved spaced-callee
   form (Replace .Left parses into the callee expression).
2. F2: Erase/Input #/Line Input #/Get # operands surface as unknown
   statements with read accesses. Fix: record write targets for those
   syntax kinds (file_number_literal excluded; Get uses TargetID only).
3. F3: VBA273 fires on ParamArray with an impossible suggestion. Fix:
   exclude ParamArray from VBA273 and VBA277.
4. F4: cfg/clone.go cloneParameters missed PassingRange. Fix: clone it and
   extend the snapshot-isolation test.
5. Follow-ups: VBA276 gains the constrained exclusion (VBE rejects
   Implements passing-modifier mismatches in both directions; verified on
   real Excel). Indexed element writes (v(0) = 1) become possiblyWritten:
   not a VBA274 reassignment, still caller-visible so VBA275 stays
   suppressed. Call Mutate((x)) is forced ByVal and no longer propagates.
   Missing mutation summary now returns nil explicitly (fail-open).

# Issue #823 final-review pass 2 follow-up

Review run_7c453f1dc6a5 pass 2 on commit 378b1333. Three confirmed findings
verified against the report's IR dumps, probe outputs, and the implicated
source; all valid. Plan:

1. F-P2-1: named-argument ExpressionIDs bind the widest expression in the
   whole `a:=x` range, so a name at least as long as its value hides the
   value from propagation (VBA275 false positive) and from every consumer
   that assumes ExpressionIDs[i] == Named[].ExpressionID
   (object_use_before_set, array_safety_source_order, object_container_flow,
   variable_assignment). Fix in procedureir/visitor.go: use
   argument.valueRange when the argument is named; add a width-inverted
   regression test for parameter passing and named-argument coverage for the
   sibling consumers.
2. F-P2-2: indexed call expressions (arr(0)) emit no receiver access, so
   element-valued arguments, With receivers, and file-statement operands are
   invisible to propagation (VBA275 false positives). Fix: resolve
   ExpressionCall roots through the callee child in parameterExprRootName,
   generalize the implicit-member argument fallback to a root-resolution
   fallback for any unhandled argument, and mark ExpressionCall write
   targets possiblyWritten in recordParameterWriteTarget.
3. F-P2-3: parameterMemberTargetRoot stops at spaces so [My Field] loses
   member writes, and parameterTypeIsObject runs before the UDT check so a
   UDT shadowing a builtin object name (e.g. Collection) drops member
   writes. Fix: extract bracketed names whole, fall back to expression
   resolution whenever the text root matches no parameter, and check UDT
   membership before the builtin object list.
4. Regression tests for each, plus docs/spec updates (named mapping is now
   width-independent; element-valued arguments and receivers covered).

# PR #849 review and security follow-up (Issue #828)

- [x] Update tree-sitter-vba v0.14.5 in the third-party licence inventory; verify the exact security gate locally.
- [x] Fix conditional-procedure scope for VB087. VB074 checks every parsed branch body; VB090 fails open for branch-dependent labels.
- [x] Make VB088 inspect only declaration visibility and cache module kind once per scan.
- [x] Cover numeric labels and their branch references for VB090.
- [x] Report VB081 on same-source declared suffixed callees while leaving unknown and intrinsic callees unreported.
- [x] Confirm focused and full Go suites, corpus snapshots, lint/docs, and security inventory locally; push a review-fix commit to PR #849.

# Issue #852 final-review pass 2 follow-up

- [x] Make template-owned PROJECT topology validation deterministic by checking removals/renames in original declaration order and additions in requested spec order.
- [x] Add regressions that assert the first reported conflict for multiple invalid components.
- [x] Run focused pack tests, formatting checks, and commit the review fix.

# PR #861 review follow-up

- [x] Validate `Attribute VB_Name` for existing standard and class component updates before assigning normalized source.
- [x] Add regression coverage for mismatched existing standard and class modules while preserving document/UserForm header behavior.
- [x] Run focused and full validation, then push the fix and reply to the review thread.

# PR #903 review follow-up

- [x] Require REFERENCECONTROL OriginalTypeLib identity rather than trusting a twiddled LIBID; cover conflicting identities and atomic rejection.
- [x] Reject incomplete generated-form expected maps before Excel starts; compare SpinButton Delay and ScrollBar Delay/ProportionalThumb with regression coverage.
- [x] Align .NET builtin control contracts with the canonical registry and fix the failing Windows/Linux CI test.
- [x] Verify focused Go/.NET/PowerShell tests, lint/docs, and staged formatting.
- [x] Rerun the fresh Excel generation gate after the concurrent session finishes: Delay/ScrollBar ProportionalThumb, both forms, save/reopen, sentinel, and normalized binary readback passed; owned PID 204972 cleanup confirmed. Evidence and command recorded in the fixture README.
- [x] Push the fixes (`4e735e72`) and reply to the three review comments with focused regression and fresh Excel evidence; verify the updated PR body on GitHub.

# Issue #886

- [x] Parse complete references and ensure first-UserForm Forms reference atomically.
- [x] Enable canonical spec/code discovery and common-control generation in blank pack.
- [x] Cover reference preservation, code modes, capability failures and publication safety.
- [x] Validate generated artifacts in Excel, retain fixtures and run release gates.
- [x] Update contracts/docs, run checks and independent review. Reference reviewer found no confirmed defects and independently passed pack/fuzz tests; blank-input reviewer found no confirmed defects. Case-variant directory spelling remains an unconfirmed follow-up, not a changed contract. Evidence: `internal/pack/vbaproject/testdata/forms-reference-excel/README.md`.

# PR #908 review follow-up

- [x] Clarify case-sensitive retained names and case-insensitive type labels.
- [x] Check the edited pack output's module identity rather than the original output.
- [x] Reproduce default TabIndex collisions during topology edits; reserve final sibling retained/explicit values before assigning generated defaults, including property-bag aliases, and preserve existing values.
- [x] Run focused regressions, affected package tests, CLI pack tests, and required lint/docs checks.

# Issue #885 MultiPage / Page / TabStrip

- [x] PR #910 review follow-up: skip observed Page geometry; retain nested omitted tabs using the inspection type resolver; defer and verify sibling TabIndex; populate lists before selection/text; prevent caption normalization from adding exported source lines. Bridge regressions pass (507 full / 59 focused); session-backed Excel build/apply/snapshot/save and reopened snapshot succeed in `tmp_workspaces/issue-910-review-e2e`.
- [x] PR #910 final-selection acceptance: the pack gate exposed copied observed selection from the old Page collection. Discard unchanged selection observations before final-topology validation and retain original observations in the property-only stage. Six template/edit regressions and related Go tests/lint pass. Pure-Go pack, Excel runtime, save/reopen runtime and confirmed cleanup pass in `tmp_workspaces/issue-910-selection-pack` and `tmp_workspaces/issue-910-selection-excel`.

- [x] Final review P2: synchronize Page Site visibility after topology reconciles the final selected identity; cover selected-page removal, additions, reorder and empty topology. Userforms/pack/filepull tests, lint, docs checks and the retained pure-Go reproduction pass. Review pass 1 report and reproduction are preserved in `tmp_workspaces/final-review/pass-1/`; review resources are cleaned up.
- [x] Final review pass 2 verified `ffb16e22579fd3931aace7ee6bca6c68fe9f257b`: no confirmed findings; compiler/oforms tests and retained reproduction pass. Report preserved in `tmp_workspaces/final-review/pass-2/`; reviewer, watcher, setup terminal and temporary worktree released/removed. Excel/COM was not rerun after the selection fix; remote CI remains unrun.
- [x] PR #910 CI follow-up: update CLI/LSP fixture assertions for the added MultiPage/Page/TabStrip controls and supported MultiPage selection; both formerly failing tests pass. Local full Go run also hit the unchanged analyzer package's default 10-minute timeout; do not report it as passed.

- [x] Capture sequential Excel-authored and saved/reopened differential evidence, including empty Pages/Tabs and layout.
- [x] Add bounded lossless TabStrip arrays/flags and MultiPage x bookkeeping parsing and writing.
- [x] Extend new generation and atomic before/after topology/property compilation, preserving opaque state.
- [x] Align FormSpec, projection, editor intelligence, and Excel authoring/inspection contracts.
- [x] Verify blank/template pack, file pull, Excel persistence/runtime, and relevant package/bridge tests.
- [x] Update MS-OFORMS/pack/form specifications, ADR-0012, user documentation and changelogs; self-review.

Evidence and exact local commands are retained in the compiler's
`testdata/multipage-excel-authored/README.md` and
`testdata/multipage-excel-generated/README.md`. Focused persistence review
confirmed and then rechecked fixes for empty cached-tab edits and quadratic
signature membership checks. Local checks passed for userforms/pack/filepull,
CLI pack/form, 51 bridge tests, Go lint, VS Code TypeScript and docs. No remote
CI or publication is claimed.

# Issue #912 final review follow-up

- [x] Validate the pass 1 P2: CLI encoding preflight misclassifies referenced pictures with source suffixes.
- [x] Resolve validated picture roles before encoding and share command-local role metadata with symbol discovery, preserving full-project diagnostics.
- [x] Add CLI pack/file-push publication regressions for .bas/.cls/.frm pictures, a code-sidecar collision, and an internal symlink; verify persisted bytes and ordinary invalid-VBA rejection.
- [x] Run affected/full Go checks, corpus snapshots, lint and fresh installed-binary Excel picture gates.
- [x] Complete analyzer performance comparison: deterministic counters unchanged, no suspicious metric increases; concurrent test activity prevents isolated timing claims. Pass 2 confirmed the first fix and found a Linux case-distinct picture/sidecar identity defect.
- [x] Reproduce the pass 2 P2 in Linux, share host path keys across inventory/preflight/symbol discovery/file-push dependencies, and verify actual CLI pack/push preserve case-distinct sidecar code and reject invalid code before publication.
- [x] Complete final full Go/lint/format checks and fresh installed-binary Excel file-push/blank-pack/template-pack picture gates after the structural path-identity correction.
- [x] Run exceptional pass 3 on 2ab038f7: prior findings resolved, one new in-scope P2 confirmed for case-distinct Assets/ nested forms. Stop the full review loop at three passes.
- [x] Preserve pass 3 evidence and clean up all review resources. Correct reserved assets comparisons using the shared host path key; Linux inventory/actual-push regressions and full Windows Go/lint/format checks pass. Fresh installed-binary Excel file-push/blank-pack/template-pack gates pass with confirmed cleanup. Final correction receives coordinator regression/self-review verification; no fourth independent pass or remote CI is claimed.

# v0.35.0 bundled agent skill refresh

- [x] Correct backend routing and obsolete UserForm restrictions in the entrypoint; retain session-first guidance for Excel work.
- [x] Refresh forms reference with canonical source authority, containers, pictures and omission semantics; add focused pack reference and architecture discovery.
- [x] Verify installed reference links across providers and execute the documented examples through blank/template pack and file push with artifact readback. Full agentskill/CLI tests PASS (1.350s/136.836s), lint and skill validator PASS; installed-binary skill smoke PASS in tmp_workspaces/issue912-agent-skill-release/xlflow.
- [x] Complete self-review and docs/lint checks; prepare the skill update for commit and resume the authorized PR preparation. Validation evidence: tmp_workspaces/final-review/agent-skill-validation.md.

# PR #928 security CI follow-up

- [x] Inspect the failed security job and confirm the fixed Go patch release against the official release and vulnerability feeds.
- [x] Raise the Go patch baseline to 1.26.9 and verify the complete security scan and Go tests (analyze: 589.162s; LSP: 223.591s).
- [x] Prepare the verified security fix for the existing PR.

# Unreleased UserForm canonical authoring and UX

- [x] Define one client-size authoring schema and update Designer/LSP/projection.
- [x] Share pure-Go generation and root resizing across pack, push and Excel reflection.
- [x] Remove mandatory prebuilt .frm artifacts for canonical specs; preserve code and opaque state.
- [x] Persist scaffold VBA codepage and add actionable encoding diagnostics / pack --check.
- [x] Update ADR/specs/user documentation and changelogs.
- [x] Run focused regressions, required checks and fresh session-backed Excel / pack round-trips.

# Deferred-save UserForm Designer

## Initial MultiPage and TabStrip members

- [x] Create two initial Pages or tabs on Toolbox placement, selecting the first member and allocating collision-free names and IDs.
- [x] Admit the narrow atomic creation transaction through Webview, host and LSP capability/validation boundaries; preserve one-step Undo/Redo and deferred source saving.
- [x] Resolve the review finding: reject standalone empty collection insertion messages and gate MultiPage/TabStrip insertion by owner type on every extension layer. Read-only second pass confirms the finding is resolved.
- [x] Full VS Code 1.141.1 tests PASS, including visible headers, collection capability matrix, native history and explicit-only source publication. Full LSP tests PASS (131.894s), covering YAML/JSON source preservation, nested Frame/Page placement and atomic rejection. Lint/docs/format and diff checks PASS. ADR/spec/guides and both changelogs updated, latest VSIX rebuilt and `rtk task install` succeeded. No Excel E2E is needed for the Toolbox/LSP-only change.

## Explicit save false conflict

- [x] Reproduce false conflicts with multiline LF edits in a native CRLF source and a real VS Code save participant before each fix.
- [x] Match native TextDocument edit normalization and reconcile successfully saved output from the published baseline while retaining pre-save external-change protection and safe draft history.
- [x] Full VS Code 1.141.1 tests PASS, including both regressions, repeated Save and Undo/Redo. Lint/docs/format and diff checks PASS; one-pass read-only review found no confirmed issues. Save-contract specification and both changelogs updated; VSIX rebuilt. No Excel E2E is needed for this editor-only fix.

## Explicit source save with Auto Save enabled

- [x] Separate internal draft checkpoints from explicit YAML/JSON publication; add Designer-scoped Ctrl+S/Cmd+S and keep global Auto Save settings untouched.
- [x] Preserve URI-shared native Undo/Redo, pending-source badge and durable draft recovery on close/reopen/restart; retain corrupt recovery data and failed saves. Read recovery records behind pending source-queue writes so immediate reopen cannot miss the final edit.
- [x] Full VS Code 1.141.1 tests PASS, including actual Auto Save, explicit command/button, direct source Auto Save, native history, normal reopen/provider restart, JSON structural conflicts, Revert, Save As and malformed-record retention. Lint/docs checks PASS; root pnpm version switching required a command-scoped environment override for the existing local package-manager change. ADR/specs/guides/changelogs updated; VSIX packaged. Bounded read-only review reported no confirmed findings; no Excel E2E is needed for this editor-only change.

- [x] Record the URI-shared draft, native custom-document history, conflict, save, and VS Code lifecycle decisions in ADR-0073; supersede the relevant ADR-0064/0066 decisions.
- [x] Update the Designer specification, user guide, and unreleased changelogs for deferred save and explicit conflict resolution.
- [x] Implement `CustomEditorProvider` and URI-shared draft state, including draft history rebasing, `CustomDocumentEditEvent` Undo/Redo, and Save/Save As/Revert/backup/restore through the latest `TextDocument`. Isolate native Designer history with a dedicated editor resource URI while retaining the real source URI for LSP and persistence.
- [x] Verify shared drafts across panels; unchanged source before Save; external-edit merge and history rebasing; structural and same-line conflict retention; explicit conflict resolution; Save As, source-preserving Revert, backup/recovery, and configured AutoSave. Full VS Code 1.141.1 tests, type checks, lint/docs checks and diff checks pass. Bounded read-only review completed; source-only change requires no Excel workbook E2E.
