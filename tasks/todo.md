# Issues #340 and #341: LSP performance phase 1

- [x] Finish #340 opt-in performance instrumentation and document-version metadata
- [x] Finish deterministic large-project/large-module benchmark fixtures
- [x] Verify #340 with Go, VS Code, lint, docs, and benchmark runs
- [x] Record #340 baseline on GitHub
- [x] Publish the #340 draft PR (#351)
- [ ] After #340 lands on main, implement #341 diagnostics coalescing on a separate branch

---

# Issue #335: Workbook recovery quarantine

- [x] Add ADR-0020 and permanent recovery-state specifications
- [x] Add versioned atomic per-workbook recovery metadata
- [x] Extend central coordination policy with recovery behavior
- [x] Enforce recovery after lease acquisition in CLI and direct Runner paths
- [x] Publish recovery before lease release for uncertain bridge outcomes
- [x] Add `recovery clear [--force]` and `session stop --discard`
- [x] Clear recovery through confirmed session discard and process cleanup
- [x] Make status/session status/process list recovery-aware without unsafe COM probes
- [x] Add VS Code Recovery Required state and safe recovery actions
- [x] Run focused and full Go, .NET, VS Code, lint, docs, and diff verification
  - `task test`, `task lint`, `pnpm docs:build`, .NET 371 tests, VS Code
    extension tests, .NET format verification, and `git diff --check` passed
  - Full tests leave 0 per-user recovery marker files
- [x] Run Windows Excel release-gate and recovery E2E under `tmp_workspaces`
  - Primary: `C:\dev\go\xlflow\tmp_workspaces\recovery-quarantine-e2e`
  - Init: `C:\dev\go\xlflow\tmp_workspaces\recovery-quarantine-init`
  - Blank scaffold, doctor, pull/lint, standard/class/UserForm + `.frx`,
    session-first run/test/save, Designer snapshot, COM sentinel, and init passed
  - Timeout published `required=true`, `published=true`, and the affected Excel
    PID before follow-up commands could enter
  - `--wait` failed immediately with `workbook_recovery_required`
  - Recovery status used metadata only; process list skipped the workbook probe
  - Managed discard, PID cleanup, verified clear, and force clear all exercised
  - Excel PIDs observed: 26836, 62108, 60560, 51048, and 49040
  - Final state: 0 Excel processes and 0 recovery markers
- [x] Complete staff-level self-review

---

# Issue #325: UserForm/Designer coordination coverage

- [x] Confirm all workbook-backed UserForm leaves use the shared workbook lock
- [x] Add exhaustive central policy and bridge-selector coverage
- [x] Add held-lock CLI and direct Runner tests that stop before handlers/bridge
- [x] Cover Designer owner metadata, bounded wait, cancellation, and timeout
- [x] Cover different-workbook independence and WSL delegation classification
- [x] Update coordination specs, user docs, agent guidance, and changelog
- [x] Run Windows Excel UserForm/Designer E2E
  - Primary workspace: `C:\dev\go\xlflow\tmp_workspaces\issue-325-e2e`
  - Independent workbook: `C:\dev\go\xlflow\tmp_workspaces\issue-325-e2e-other`
  - Installed the paired Go CLI and .NET bridge with `task install`
  - Exercised JSON commands `new`, `form new`, `session start`, `form build`,
    `push --fast --session --no-save`, `save --session`, `form snapshot`,
    `inspect form --designer`, `form build --overwrite`, `form export-image`,
    `run`, `pull`, `lint`, and `session stop`
  - Initial push before the first build correctly failed preflight because the
    new spec had no generated `.frm`; build created the artifacts and the
    subsequent push succeeded
  - Build/snapshot/inspect observed caption `Coordination Form`, two controls,
    width 240, and height 150; overwrite succeeded and PNG export produced a
    382x241 image
  - `.frm` / `.frx` survived save/stop/pull at 603 / 2584 bytes, and pull
    exported one sidecar code-behind file; lint reported no issues
  - During `run Main.HoldWorkbook --session`, status reported busy owner
    `run` / `execute`; immediate `form snapshot` returned exit 3
    `workbook_busy`, operation `form.snapshot`, and created no output file
  - `--wait --wait-timeout 15s form snapshot` succeeded after release and
    created `artifacts/waited.yaml`; the macro wrote and live Excel inspection
    observed `Sheet1!A1 = coordination-finished`
  - While the primary workbook was busy, snapshotting `OtherForm` in the
    independent workbook succeeded and created `artifacts/other.yaml`
  - Not repeated on a physical WSL host; WSL source/workbook delegation is
    covered by the expanded command table. Actual simultaneous Designer COM
    calls were replaced with deterministic held-lease Designer-owner tests.
- [x] Run full tests, lint, docs build, and diff check
  - `task test`, `task lint`, `pnpm docs:build`, and `git diff --check` passed
- [x] Complete final staff-level diff review
  - No High/Medium findings; restored explicit run/test policy assertions and
    widened new wait-test timing margins from the two Low findings
- [ ] Create, review, and merge PR #325

---

# Issue #324: Session status coordination

- [x] Probe the canonical configured workbook lock before bridge status
- [x] Add the additive top-level coordination payload and fallback warning
- [x] Preserve existing success/failure session envelopes and warnings
- [x] Render idle, busy, ownerless, and public owner details
- [x] Add focused payload, ordering, warning, JSON, and human-output tests
- [x] Run full tests, lint, docs build, and Windows Excel E2E
  - Workspace: `C:\dev\go\xlflow\tmp_workspaces\issue-324-e2e`
  - During `run Main.HoldCoordination --session`, status reported busy owner
    `run` / `execute`; after completion it reported `busy:false`
  - Repeated without `--session`; run reported `session.mode=auto`, status
    reported the same busy owner, and returned to idle after completion
  - Live `Sheet1!A1` observed `coordination-finished`; session saved and stopped
- [x] Create, review, and merge PR #324

---

# Issue #322: Optional workbook wait and timeout

- [x] Define the opt-in 30-second wait contract and ADR
- [x] Add global wait options and policy-based validation
- [x] Add total acquisition timeout, cancellation, and partial lease cleanup
- [x] Keep JSON output progress-free and human wait output concise
- [x] Add focused wait, timeout, cancellation, and multi-target tests
- [x] Verify WSL argument forwarding
- [x] Run full tests, lint, docs build, and self-review
- [x] Create, review, and merge PR #322

---

# Issues #320, #321, #323: Workbook operation coordination

## Implementation

- [x] Add cross-process workbook lock and crash-safe lease lifecycle
- [x] Integrate exclusive coordination across full workbook command lifetime
- [x] Add structured `workbook_busy` diagnostics and stable exit behavior
- [x] Add atomic owner metadata with stale-owner race protection
- [x] Update ADR, specs, public docs, and changelog
- [x] Add subprocess contention, crash, diagnostic, and metadata tests
- [x] Run focused and full Go verification
- [x] Run Windows + Excel COM E2E in `tmp_workspaces`
- [x] Self-review the complete diff

---

# Issue #79 + Issue #78 follow-up

---

# Issue #82: Windows release packaging for .NET bridge

- [x] Update `scripts/build-dotnet-bridge.ps1` to `dotnet publish` self-contained single-file output
- [x] Wire GoReleaser to build/package `xlflow-excel-bridge.exe`
- [x] Add `.NET` setup and bridge tests to release workflow
- [x] Update README / README.ja / VitePress / bridge docs for Windows sidecar packaging and policy caveats
- [x] Run release-oriented verification (`goreleaser check`, snapshot build, archive inspection)
- [x] Verify extracted Windows archive resolves the sidecar bridge with `doctor --bridge dotnet --json`

- [x] Add ADR for hybrid UIA + Win32 dialog correlation and child worker isolation
- [x] Implement reusable .NET DialogWatcher subsystem
- [x] Add dialog snapshot schema, fingerprints, and safe action policy
- [x] Add disposable .NET macro worker and parent-side timeout management
- [x] Integrate runtime and compile dialog diagnostics into `.NET run`
- [ ] Restore `.NET run` runtime markers, trace, direct/fast, and stream parity
- [x] Add `double` typed arguments to Go, PowerShell, and .NET contracts
- [x] Update public specs, bridge docs, and changelog
- [x] Run Go and .NET tests
- [ ] Run Windows Excel COM E2E with fresh `tmp_workspaces`

---

# Bug Fix: XlflowDebug ParamArray + run compile watcher

## PR #92 CI/security and .NET bridge review

- [x] Inspect PR #92 review comments and identify actionable unresolved thread
- [x] Confirm .NET pull/push service implementations exist under `bridge/dotnet/src`
- [x] Run local .NET bridge tests to verify registered pull/push commands compile
- [x] Update `go.mod` to Go `1.26.4` for standard-library vulnerability fixes
- [x] Add .NET bridge tests to the CI test job with `setup-dotnet`
- [x] Run full Go, .NET, vet, and govulncheck verification after edits
- [x] Fix CI-only .NET bridge failure caused by untracked `bridge/dotnet/src` pull/push files
- [x] Anchor `.gitignore` project source rule as `/src/`
- [x] Re-run `go test -count=1 ./internal/excel/bridge`, .NET bridge tests, and `go vet ./...`

## Welcome header refresh for new/init

- [ ] Update `internal/cli/welcome.go` to render `Welcome to` above the logo
- [ ] Add command reference URL below the logo and above version text
- [ ] Mute the `Version`/URL info color slightly compared to the current welcome info text
- [ ] Update `internal/cli/welcome_test.go` expectations for content and ordering
- [ ] Update `internal/cli/root_test.go` interactive welcome expectations and add `new` regression coverage if needed
- [ ] Run focused `internal/cli` tests for welcome rendering and command output
- [ ] Add an Unreleased changelog note for the welcome screen refresh

## Status

- [x] Investigated `ParamArray` forwarding issue in `defaultDebugRuntimeModule`
- [x] Removed `JoinLogMessage(Parts)` forwarding and built the message inline in `Log` — Fix A
- [x] Investigated `Invoke-XlflowVBECompile` catch block missing `$result.ok = $false` — Fix B
- [x] Added `$result.ok = $false` in common.ps1 catch block for compile control lookup failure
- [x] Added regression test: `TestXlflowDebugLogDoesNotForwardParamArrayToHelper` (`scaffold_test.go`)
- [x] Added regression test: `TestInvokeXlflowVBECompileMarksFailureWhenCompileControlNotFound` (`scripts_test.go`)
- [x] Updated CHANGELOG.md with both fixes under Unreleased
- [x] Updated tasks/feature_spec.md with bug fix spec
- [x] Updated tasks/lessons.md with ParamArray forwarding lesson
- [x] All scaffold tests pass (28)
- [x] All compile/dialog/run tests pass (5+20 script parse)
- [x] go vet clean
- [x] PS lint passes (22 files)

## E2E Verification

- [x] `xlflow-tmp-workspace-e2e` skill: session-first workflow with `SampleFail` macro
- [x] Confirm `ParamArray` compile error does not occur — `SampleFail` produces `macro_failed` (runtime type mismatch `HRESULT 0x800A9C68`), no `XlflowDebug.bas` ParamArray error
- [x] Confirm no GUI dialog residual — all commands return normally, session save/stop clean
- [x] Confirm structured failure in terminal output — `status: failed`, `error.code: macro_failed`, `phase: invoke_macro`
- [x] Standard pull/lint flow passes cleanly after session stop
- [ ] Fix B compile control not-found path — E2E verification not possible in this environment (compile control is available); regression test `TestInvokeXlflowVBECompileMarksFailureWhenCompileControlNotFound` passes

## E2E workspace

- `C:\dev\go\takt-worktrees\20260524T0200-xlflow-issue-bug-high-task-bri\tmp_workspaces\paramarray-e2e` — Fix A/B E2E verification

## Runtime diagnostic follow-up

- [x] Reproduced live break-mode selection on `Main.SampleFail` and confirmed the real failing statement is `Main` line 9 (`x = "abc"`) in `tmp_workspaces\runtime-diagnostic-e2e`
- [x] Added runtime selection scoring so user-code statements beat structural lines and temporary `XlflowRun_*` harness panes
- [x] Moved macro execution for dialog-watched runtime runs into a disposable child `powershell.exe` process so parent CLI can recover even if `excel.Run(...)` remains blocked after break-mode entry
- [x] Verified `Get-XlflowRuntimeDebugSelectionByProcessId` captures `Main:9` and resets break mode on the live Excel instance
- [ ] Full `xlflow run --session --diagnostic` E2E remains blocked in this environment by the separate pre-existing `VBE Compile command was not found.` compile-gate issue before runtime invocation

## Compile gate follow-up

- [x] Inspected real VBE command bars and confirmed `Compile VBAProject` is exposed under `CommandBars("Menu Bar") -> Debug` as control `Id = 578`, not on the `CommandBars("Debug")` toolbar
- [x] Updated compile control lookup to search by control id and menu-bar Debug popup before toolbar fallbacks
- [x] Treated `compile control exists but Enabled = false` as "already compiled / no compile needed" instead of a hard failure
- [x] Added regression tests for `FindControl` lookup, menu-bar fallback, and disabled compile control handling
- [x] Verified `Invoke-XlflowVBECompile` returns `ok=true` on the live session workbook in `tmp_workspaces\runtime-diagnostic-e2e`
- [x] Verified full `xlflow run Main.SampleFail --session --diagnostic --json` now returns structured `macro_failed` with `source=Main`, `line=9`, and nearby code for `x = "abc"`

---

# xlflow status implementation

## Status

- [x] Read existing plan (`.takt/runs/20260523-071517-issue-26-github-issue-26-xlflo/reports/plan.md`)
- [x] Add `xlflow status` CLI command (`internal/cli/root.go:2144`)
- [x] Add `xlflow status --json` with stable JSON envelope
- [x] Implement `buildStatusProject()` — project root, workbook path, src_paths
- [x] Implement `buildStatusState()` — source/workbook freshness, mtimes, push state mtime
- [x] Implement `buildStatusSession()` — status-dedicated session probe with `running`, `workbook_open`, `metadata` from session.ps1
- [x] Implement `buildStatusWarningsAndHints()` — actionable warnings/hints for dirty session, source newer, live newer
- [x] Add human-readable renderer (`renderStatus()`) with section headers (Project/Session/State/Hints)
- [x] Remove dead code (`boolValueOKForStatus` duplicate, `running`/`workbook_open` rendering in renderStatus)
- [x] Fix P0: `live_session_newer_than_disk` dead code — statePayload derived from session
- [x] Fix P0: `inspectStateForWorkbook` warnings not recycled into status
- [x] Fix P0: Contradictory state/session `source_of_truth`/`workbook_saved`
- [x] Fix: `inspectStateForWorkbook` no longer leaks `running`/`workbook_open`/`metadata` into inspect output
- [x] Fix: Unused `project` param removed from `buildStatusWarningsAndHints`
- [x] Add CLI tests: `TestStatusJSONBaseline`, `TestStatusJSONSourceNewerThanWorkbook`, `TestStatusJSONSessionFieldShape`, `TestStatusWarningsExcludeInspectSpecificMessages`
- [x] Add unit tests: `TestBuildStatusWarningsAndHintsSessionDirty`, `TestBuildStatusWarningsAndHintsSourceNewer`, `TestBuildStatusWarningsAndHintsSessionInactive`, `TestBuildStatusWarningsAndHintsProducesStatusSpecificCodes`
- [x] Add output tests: `TestWriteWithOptionsRendersStatusBaseline`, `TestWriteWithOptionsRendersStatusSessionActiveDirty`, `TestWriteWithOptionsRendersStatusSourceNewerThanWorkbook`, `TestWriteWithOptionsRendersStatusSessionLiveNewerThanDisk`, `TestWriteWithOptionsRendersStatusSectionHeaders`
- [x] Add test: `TestInspectStateForWorkbookExcludesStatusOnlyFields` (regression guard for inspect contract)
- [x] Update `docs/specs/cli-contract.md` — status JSON contract with heuristic note
- [x] Create `vitepress/commands/status.md` — with `running`/`workbook_open`/`metadata` in session example
- [x] Update `vitepress/commands/index.md` — add status row
- [x] Update `vitepress/ai-agents/index.md` — add status to agent loops
- [x] Update `vitepress/concepts/workbook-session-source.md` — add status usage
- [x] Update `vitepress/concepts/source-of-truth.md` — add status usage
- [x] Update `README.md` — add status command
- [x] Update `README.ja.md` — add status command
- [x] Update `CHANGELOG.md` — add status entry
- [x] Full CLI test suite passes (10.4s)
- [x] Full output test suite passes (0.6s)
- [x] `go vet` clean
- [x] Windows + Excel COM E2E — `session start -> push --no-save -> status` confirms all warnings/hints fire, session payload has 14 fields
- [x] Windows + Excel COM E2E — `save -> status` confirms divergence resolved (0 warnings/hints)
- [x] Windows + Excel COM E2E — `inspect workbook` session payload correctly has 11 fields (no `running`/`workbook_open`/`metadata`)

## E2E workspaces

- `C:\dev\go\xlflow\tmp_workspaces\status-e2e` — initial E2E
- `C:\dev\go\xlflow\tmp_workspaces\status-e2e-fix` — P0 fixes verified
- `C:\dev\go\xlflow\tmp_workspaces\status-e2e-v2` — session field fix verified
- `C:\dev\go\xlflow\tmp_workspaces\status-e2e-v3` — separation of concerns verified (status 14 fields, inspect 11 fields)

## Unverified

- None

---

# Takt workflow monitor alignment

## Todo

- [x] Inspect existing `.takt` workflow and persona patterns
- [x] Define minimal `self_review -> fix -> self_review` loop
- [x] Add `.takt/facets/personas/supervisor.md`
- [x] Add `fix` step and reroute `self_review` to it
- [x] Validate YAML/persona references and summarize residual risks
- [x] Mirror the same loop-monitor structure into `.takt/workflows/xlflow-orchestra-low.yaml`

---

# `xlflow fmt` write_tests

## Todo

- [x] Re-read `order.md`, policy, lessons, and current CLI/output tests
- [x] Add failing CLI contract tests for top-level `fmt` registration and flags
- [x] Add failing output rendering tests for human-readable `fmt` summaries
- [x] Run focused `go test` for touched packages and capture failures as implementation targets

# `xlflow fmt` implementation

## Todo

- [x] Create `internal/vbafmt/` package: `fmt.go` (formatter + file discovery + diff), `fmt_test.go` (27 tests)
- [x] Add `fmtCommand()` to `internal/cli/root.go` with `--write`/`--check`/`--diff`/`--stdin` flags
- [x] Register `fmt` in `rootCommand()` AddCommand list
- [x] Add `renderFmt()` to `internal/output/output.go`
- [x] Run focused tests: `go test ./internal/vbafmt ./internal/cli ./internal/output` — all pass
- [x] `go vet` clean on affected packages
- [x] Pre-existing write_tests tests (`TestRootCommandIncludesFmtCommand`, `TestWriteWithOptionsRendersFmtSummary`) now pass

---

# .NET doctor implementation

## Status

- [x] .NET `DoctorCommand` implemented: OS, architecture, runtime, Excel COM activation, version/build, VBIDE access, AutomationSecurity, Trust VBA access
- [x] `ExcelDiagnostics.Probe()` returns structured probe result with COM error mapping
- [x] `BridgeError` uses `HResult` field; .NET `JsonNamingPolicy.SnakeCaseLower` serializes as `h_result` matching the public envelope
- [x] Go-side `ScriptResult.Diagnostics` passes through `diagnostics` from .NET bridge response to `output.Envelope.Diagnostics`
- [x] Protocol version mismatch check preserved: `bridge_protocol_mismatch` error when `.NET` returns wrong version
- [x] `--bridge dotnet` strict mode: no PowerShell fallback on failure
- [x] Human-readable `renderDoctor` handles both flat (PowerShell) and nested (.NET) diagnostics shapes via `doctorBool`
- [x] Regression tests: `TestDotNetProviderExecuteDoctorWithRealBridge`, `TestRunnerDoctorPreservesDotNetBridgeMetadataAndDiagnostics`, `TestRunnerDoctorPreservesDotNetBridgeStructuredErrorMetadata`, `TestRunnerRejectsDotNetProtocolMismatch`
- [x] .NET bridge tests cover doctor warning passthrough and bridge-host `excel` field serialization
- [x] Output tests: `TestWriteWithOptionsRendersDoctorChecklist`, `TestWriteWithOptionsRendersDoctorChecklistFromDotNetBridge`

## Docs updated

- [x] `docs/bridge/dotnet-bridge.md` — updated from "planned" to implemented, added diagnostics shape and field descriptions
- [x] `vitepress/commands/doctor.md` — replaced old `checks` example with current `diagnostics` JSON shape
- [x] `docs/bridge/bridge-protocol.md` — fixed `hresult` to `h_result` in error example to match actual .NET bridge serialization
- [x] `docs/specs/cli-contract.md` / `vitepress/commands/doctor.md` — clarified provider-specific top-level `bridge` metadata vs nested `diagnostics`
- [x] `tasks/feature_spec.md` — added .NET doctor success/failure contract section
- [x] `CHANGELOG.md` — documents native `.NET doctor` diagnostics and structured COM error output

## Verification

- [x] `dotnet test bridge/dotnet/Xlflow.ExcelBridge.sln --no-restore`
- [x] `go test ./internal/excel/...` 相当の focused run: `internal/excel/bridge_test.go`, `internal/excel/bridge/dotnet_test.go`
- [x] Windows + Excel 実機: `example/calendar-pick` で `go run ../../cmd/xlflow --json --bridge dotnet doctor`
- [x] 実機結果: `selected_bridge=dotnet`, `protocol_version=1`, `bridge.name=xlflow-excel-bridge`, `bridge.commit=dev`, `excel.com_activation=true`, `excel.vbide_access=true`, `automation_security=1`
- [x] 実機観測: `excel.version` / `excel.build` / `trust_vba_access` はこの環境では `null`

---

# Issue 80 independent verification

## Correctness fixes

- [x] Roll back partial .NET runtime defined-name injection and abort execution on injection failure.
- [x] Validate `trace disable` source-match safety before mutating the workbook.
- [x] Reject unsupported trace actions and missing workbook paths at the .NET command boundary.
- [x] Add .NET regression tests for trace command validation and source removal decisions.

## Verification

- [x] `task install`
- [x] `dotnet test bridge/dotnet/Xlflow.ExcelBridge.sln --no-restore` - 157 passed
- [x] `go test ./... -count=1` - 862 passed
- [x] `go vet ./...`
- [x] `git diff --check`
- [x] Windows + Excel issue 80 E2E workspace: `C:\dev\go\takt-worktrees\20260603T1401-issue-80-tasuku-gooru-issue-80\tmp_workspaces\issue80-independent-e2e`
- [x] `trace enable/status/disable --bridge dotnet --json`
- [x] Modified trace source failure preserves `workbook_injected=true`
- [x] Session-first `push --fast --session --no-save -> test/run --bridge dotnet -> save -> session stop`
- [x] Runtime mode, MsgBox/InputBox/FileDialog responses, UI events, debug events, save-required state, and Excel COM workbook readback
- [x] Init workspace: `C:\dev\go\takt-worktrees\20260603T1401-issue-80-tasuku-gooru-issue-80\tmp_workspaces\issue80-independent-init`
- [x] Class/UserForm round-trip workspace: `C:\dev\go\takt-worktrees\20260603T1401-issue-80-tasuku-gooru-issue-80\tmp_workspaces\issue80-release-gate-roundtrip`

## Notes

- UserForm `.frx` remained present through repeated push/save/pull cycles, but its SHA-256 changed on each Excel export; byte-identical `.frx` output is not asserted.
