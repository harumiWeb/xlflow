# Bug Fix: XlflowDebug ParamArray + run compile watcher

## Status

- [x] Investigated `ParamArray` ByRef issue in `defaultDebugRuntimeModule` (`scaffold.go:1228`)
- [x] Changed `JoinLogMessage(ByRef Parts() As Variant)` → `ByVal` — Fix A
- [x] Investigated `Invoke-XlflowVBECompile` catch block missing `$result.ok = $false` — Fix B
- [x] Added `$result.ok = $false` in common.ps1 catch block for compile control lookup failure
- [x] Added regression test: `TestXlflowDebugJoinLogMessageDoesNotForceParamArrayToByRef` (`scaffold_test.go`)
- [x] Added regression test: `TestInvokeXlflowVBECompileMarksFailureWhenCompileControlNotFound` (`scripts_test.go`)
- [x] Updated CHANGELOG.md with both fixes under Unreleased
- [x] Updated tasks/feature_spec.md with bug fix spec
- [x] Updated tasks/lessons.md with ParamArray/ByRef lesson
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
