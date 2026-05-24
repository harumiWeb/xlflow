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
