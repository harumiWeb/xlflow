# MVP Quality Hardening Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the `xlflow` MVP safer to extend by standardizing real-workbook verification, closing current regression gaps, and aligning local workflow docs with actual verification practice.

**Architecture:** Keep the current Go CLI plus PowerShell bridge architecture intact. Strengthen the system by adding focused regression tests around PowerShell transformations, documenting the standard `tmp_workspaces` verification baseline, and adding a single obvious local verification entry point rather than redesigning command behavior.

**Tech Stack:** Go, Cobra CLI, PowerShell, Excel COM, TOML, repo-local skills/docs

---

### Task 1: Align CLI contract and quality-hardening docs

**Files:**
- Modify: `C:\dev\go\xlflow\docs\specs\cli-contract.md`
- Modify: `C:\dev\go\xlflow\README.md`
- Read: `C:\dev\go\xlflow\docs\specs\mvp-quality-hardening.md`

- [ ] **Step 1: Review current contract and README against actual verified flows**

Run:

```powershell
Get-Content -Raw C:\dev\go\xlflow\docs\specs\cli-contract.md
Get-Content -Raw C:\dev\go\xlflow\README.md
```

Expected: identify any missing mention of class/userform round-trips, `tmp_workspaces` verification, or current JSON/log behavior.

- [ ] **Step 2: Update the contract only where implementation behavior has been verified**

Edit `C:\dev\go\xlflow\docs\specs\cli-contract.md` so it explicitly reflects verified MVP behavior, especially:

```markdown
- `pull` exports standard modules, classes, forms, and workbook document modules.
- userforms may produce both `.frm` and `.frx` artifacts.
- document modules are exported in source form suitable for linting and re-import.
```

- [ ] **Step 3: Add a short contributor verification section to the README**

Add concise guidance to `C:\dev\go\xlflow\README.md` covering:

```markdown
## Verification

Use the repo-local `xlflow-tmp-workspace-e2e` skill and create disposable projects under `tmp_workspaces` for real Excel COM validation.

Baseline checks:

- `xlflow new --json`
- `xlflow doctor --json`
- `xlflow pull --json`
- `xlflow lint --json`
- module/class/form round-trip with `push/run/pull/lint`
```

- [ ] **Step 4: Review the edited docs for consistency**

Run:

```powershell
Get-Content -Raw C:\dev\go\xlflow\docs\specs\cli-contract.md
Get-Content -Raw C:\dev\go\xlflow\README.md
```

Expected: wording matches actual verified behavior and does not promise new commands or unsupported automation.

- [ ] **Step 5: Commit**

```powershell
git add C:\dev\go\xlflow\docs\specs\cli-contract.md C:\dev\go\xlflow\README.md
git commit -m "docs: define mvp verification baseline"
```

### Task 2: Add regression tests for form/class/document-module round-trips

**Files:**
- Modify: `C:\dev\go\xlflow\scripts\scripts_test.go`
- Modify: `C:\dev\go\xlflow\scripts\common.ps1`
- Modify: `C:\dev\go\xlflow\scripts\pull.ps1`
- Test: `C:\dev\go\xlflow\scripts\scripts_test.go`

- [ ] **Step 1: Write failing tests for remaining risky transformations**

Add tests in `C:\dev\go\xlflow\scripts\scripts_test.go` that cover:

```go
func TestDocumentModuleContentDropsClassHeaderLines(t *testing.T) {}
func TestUserFormExportRetainsFrxReference(t *testing.T) {}
func TestNormalizeDocumentModuleFileKeepsExecutableBodyOnly(t *testing.T) {}
```

Use PowerShell snippets that exercise `Get-XlflowDocumentModuleContent` and representative `.frm` content rather than mocking behavior in Go.

- [ ] **Step 2: Run the script-focused tests to verify the new cases fail if behavior is missing**

Run:

```powershell
go test ./scripts -run "DocumentModule|UserForm" -count=1
```

Expected: at least one new test fails before implementation if the behavior is not already protected.

- [ ] **Step 3: Implement the minimal PowerShell fixes or helpers needed by the tests**

Keep edits focused on:

```powershell
# C:\dev\go\xlflow\scripts\common.ps1
function Get-XlflowDocumentModuleContent { ... }
function Normalize-XlflowDocumentModuleFile { ... }
```

Do not redesign import/export flow. Only make the transform and helper behavior match the verified source contract.

- [ ] **Step 4: Re-run the focused script tests**

Run:

```powershell
go test ./scripts -run "DocumentModule|UserForm" -count=1
```

Expected: PASS.

- [ ] **Step 5: Run the full test suite**

Run:

```powershell
go test ./...
```

Expected: PASS with no new package failures.

- [ ] **Step 6: Commit**

```powershell
git add C:\dev\go\xlflow\scripts\common.ps1 C:\dev\go\xlflow\scripts\pull.ps1 C:\dev\go\xlflow\scripts\scripts_test.go
git commit -m "test: harden workbook roundtrip regressions"
```

### Task 3: Add one obvious local verification entry point

**Files:**
- Modify: `C:\dev\go\xlflow\Taskfile.yml`
- Modify: `C:\dev\go\xlflow\README.md`
- Read: `C:\dev\go\xlflow\.agents\skills\xlflow-tmp-workspace-e2e\SKILL.md`

- [ ] **Step 1: Inspect current task runner configuration**

Run:

```powershell
Get-Content -Raw C:\dev\go\xlflow\Taskfile.yml
```

Expected: identify whether a verification target already exists and whether it can be extended without breaking current usage.

- [ ] **Step 2: Add a lightweight quality target**

Edit `C:\dev\go\xlflow\Taskfile.yml` to add one obvious local check such as:

```yaml
  verify:
    desc: Run the fast local verification suite
    cmds:
      - go test ./...
```

Keep this target fast and deterministic. Do not try to launch Excel COM automatically inside the Taskfile unless the repository already does that safely.

- [ ] **Step 3: Document what the task does and does not cover**

Update `C:\dev\go\xlflow\README.md` so the verification section distinguishes:

```markdown
- `task verify`: fast automated checks
- `tmp_workspaces` E2E flow: real Excel COM verification
```

- [ ] **Step 4: Run the task**

Run:

```powershell
task verify
```

Expected: PASS, or the equivalent success output from the configured task runner.

- [ ] **Step 5: Commit**

```powershell
git add C:\dev\go\xlflow\Taskfile.yml C:\dev\go\xlflow\README.md
git commit -m "build: add local verification entry point"
```

### Task 4: Refresh task-tracking and lessons after implementation

**Files:**
- Modify: `C:\dev\go\xlflow\tasks\feature_spec.md`
- Modify: `C:\dev\go\xlflow\tasks\todo.md`
- Modify: `C:\dev\go\xlflow\tasks\lessons.md`

- [ ] **Step 1: Update the session feature spec to match the active quality-hardening slice**

Write the current focus into `C:\dev\go\xlflow\tasks\feature_spec.md` with:

```markdown
# xlflow MVP Quality Hardening Spec
```

Summarize the specific slice being executed in the current session.

- [ ] **Step 2: Update the todo list with completed and pending verification tasks**

Edit `C:\dev\go\xlflow\tasks\todo.md` so it tracks:

```markdown
- automated regression coverage
- docs alignment
- local verification command
- remaining manual Excel COM checks
```

- [ ] **Step 3: Record any newly learned failure pattern**

If implementation uncovers a new durable rule, append one line to `C:\dev\go\xlflow\tasks\lessons.md`.

If no new rule was learned, leave `tasks/lessons.md` unchanged.

- [ ] **Step 4: Review task-tracking files for stale MVP-era content**

Run:

```powershell
Get-Content -Raw C:\dev\go\xlflow\tasks\feature_spec.md
Get-Content -Raw C:\dev\go\xlflow\tasks\todo.md
Get-Content -Raw C:\dev\go\xlflow\tasks\lessons.md
```

Expected: current session documents match the active work and do not point to already-finished older tasks.

- [ ] **Step 5: Commit**

```powershell
git add C:\dev\go\xlflow\tasks\feature_spec.md C:\dev\go\xlflow\tasks\todo.md C:\dev\go\xlflow\tasks\lessons.md
git commit -m "chore: refresh quality hardening task tracking"
```

## Self-Review

- Spec coverage:
  - `docs/specs/mvp-quality-hardening.md` workstreams map to Task 1 through Task 4.
  - E2E baseline and contributor workflow are covered by Task 1 and Task 3.
  - regression protection is covered by Task 2.
  - task-tracking and future agent continuity are covered by Task 4.
- Placeholder scan:
  - no `TODO`, `TBD`, or “similar to above” placeholders remain in actionable steps.
- Type consistency:
  - all file paths and command names match the current repository structure and existing command names.
