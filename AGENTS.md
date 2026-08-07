# Guide for AI Agents

## 0. Project Overview

```txt
root: .
├── bridge/
│   └── dotnet/ # Code for the .NET bridge component
├── cmd/
├── docs/ # ADR documents, specification documentation, and other development materials
├── internal/ # Internal packages and components
├── scripts/ # Automation scripts
├── tasks/ # Task management and learning records
├── vitepress/ # User documentation
├── .editorconfig
├── .goreleaser.yaml
├── AGENTS.md
├── CHANGELOG.md
├── CLAUDE.md
├── CONTRIBUTING.md
├── global.json
├── go.mod
├── go.sum
├── lefthook.yml
├── LICENSE
├── package.json
├── pnpm-lock.yaml
├── pnpm-workspace.yaml
├── PSScriptAnalyzerSettings.psd1
├── README.ja.md
├── README.md
├── SECURITY.md
├── Taskfile.yml
└── THIRD_PARTY_LICENCES.md
```

## 1. Workflow Design

### 1. Basic Approach: Work in Plan Mode First

- For tasks involving three or more steps or those affecting the overall architecture, always begin in Plan mode.
- If progress stalls at any point, do not force continuation - stop and replan instead.
- Use the Plan mode not only for implementation but also for designing verification procedures.
- As early as possible, refine specifications to reduce ambiguity.

### 2. Multi-Agent Strategy

- Make active use of subagents to avoid contaminating the main context.
- Delegate tasks such as research, verification, and parallel analysis to subagents.
- For complex problems, utilize subagents even when they require significant computational resources.
- Assign each subagent only one task to maintain focused execution.
- Use an explorer for codebase exploration (primarily reading activities).
- Use a worker for implementation and modifications.
- Use a reviewer for code reviews.

### 3. Self-Improvement Loop

- When receiving correction instructions from users, document these patterns in `tasks/lessons.md`.
- Formulate clear rules for yourself to prevent repeating the same mistakes.
- Continuously refine these rules until error rates decrease significantly.
- At the beginning of each session, review relevant lessons related to the project.

### 4. Always Verify Before Finalizing

- Do not mark tasks as complete until you can demonstrate their functionality.
- When necessary, compare your changes against the main branch for verification.
- Always ask yourself: "Would a staff engineer approve this?"
- Complete the process by running tests, reviewing logs, and demonstrating proper operation.
- For pre-release validation, do not assume CI alone is sufficient. If Windows + Excel integration requires actual E2E testing, use the `xlflow-tmp-workspace-e2e` skill to perform release verification through `tmp_workspaces`.
- When performing multiple `push`/`run`/`test`/`pull`/`save` operations on Windows+Excel real device validation, avoid non-session single-command sequences. Instead, use the basic workflow pattern: `session start -> push --fast --session --no-save -> run/test --session -> save --session -> session stop`. Reopening the workbook each time can significantly slow down verification or make it appear frozen during waiting periods.

### 5. Maintain Balance While Pursuing More Elegant Solutions

- Before implementing major changes, pause to first consider: "Is there a more elegant way to do this?"
- If your fix feels ad hoc, reframe it as: "How can I implement this in a more refined manner based on what I know now?"
- However, do not overthink simple and obvious fixes - avoid excessive design.
- Before delivering any deliverable, thoroughly review your own implementation with a critical eye.

### 6. Handle Bug Fixes Autonomously

- When receiving bug reports, investigate them independently without waiting for instructions, then proceed directly to resolution.
- Use logs, errors, and failing tests to autonomously identify and resolve the issue.
- Avoid forcing users into unnecessary context switches.

## - Even without explicit instructions, if the CI pipeline is down, take initiative to resolve it.

## 2. Required Workflow Procedures

Before generating or modifying code, perform the following steps according to the scale of your work:

1. Understand the requirements: Review relevant specification documents, ADR documentation, and existing implementations.
2. Consider the design implications: Assess impact scope, compatibility with current designs, and alternative approaches.
3. If necessary, create working notes:

- For recurrence prevention: `tasks/lessons.md`

4. Add or update tests as needed.
5. Implement changes.
6. Verify functionality.
7. Run tests.
8. Conduct self-review.
9. Update documentation, ADR documents, specifications, and the CHANGELOG as appropriate.

- Any updates to ADR documents or specifications must be recorded in the respective directories:
- For ADR documents: `docs/adr/`
- For specification documents: `docs/specs/`
- If changes affect public APIs, they may require recording in the following documentation:
- Specification documents within `docs/specs/`
- User documentation in `vitepress/`
- Overview descriptions in the `README.md` file
- When making changes that impact users, append updates to the `CHANGELOG.md` file.
- For VSCode extension modifications, append updates to `editors\vscode\CHANGELOG.md`.

### Pre-Release E2E Testing

- Before releasing any changes involving Windows + Excel COM / VBIDE access, perform real device E2E testing using the repo-local `xlflow-tmp-workspace-e2e` skill.
- Ensure you verify at least the following scenarios: blank workbook, standard module round-trip, class module round-trip, UserForm + `.frx` round-trip, and `init` functionality paths.
- For releases including the `pack` command, also perform pack artifact smoke testing by opening the generated `.xlsm` in actual Excel to compile/run minimal macros and observe observable effects like sentinel cell values (procedures can be found in Section 7 of the "Release-gate Excel smoke" section and in the `xlflow-tmp-workspace-e2e` skill documentation under `docs/specs/pack-command.md`). Restrict automated PR CI to Linux/pure Go environments only.
- If modifying session-aware workflows, also include the sequence `session start -> push --fast --session --no-save -> run/test --session -> save --session -> session stop` in the release gate checks.
- When combining multiple workbook-backed commands during real device E2E testing for Windows + Excel, even if not altering the session-aware workflow, prioritize first performing `session start -> push --fast --session --no-save -> run/test --session -> save --session -> session stop`.

### Local VBE oracle for static-analysis changes

When changing static-analysis semantics, parser interpretation used by
diagnostics, type/call/argument validation, object or `Set` diagnostics,
diagnostic severity, or LSP projections, determine whether the behavior
depends on actual VBE semantics. If it does, run the focused local VBE oracle
cases on Windows with Excel and trusted VBIDE access enabled. Add an `observe`
fixture when needed, run the known accept/reject controls first, and promote
only confirmed outcomes through the explicit promotion command. Oracle cases
must run sequentially. A timeout, unknown modal, failed Compile invocation,
worker/COM failure, or unconfirmed Excel cleanup is an infrastructure failure
and a stop-the-line condition: do not promote fixtures or change analyzer
behavior based on that result. Record the executed case IDs in the PR.

The oracle is developer-only and local. Never invoke it from production xlflow
commands, ordinary tests, or GitHub Actions. See
`docs/specs/vbe-oracle.md` for the complete workflow and contract.

When binding a compile-error diagnostic to oracle evidence, first search the
existing fixtures and use confirmed VBA behavior rather than general-language
assumptions. Bind both sides of the VBE boundary: rejected fixtures must carry
expected diagnostics and `analysis.negative_controls` pointing to accepted
fixtures whose forbidden contracts protect the same rule on every declared
surface. Do not change VBE evidence merely to make analyzer tests pass.

Run the known controls and focused cases sequentially. A timeout, unknown modal,
failed Compile invocation, worker/COM failure, or unconfirmed cleanup is an
infrastructure failure and a stop-the-line condition; do not promote evidence
or change analyzer behavior after such a result. Report executed oracle case
IDs and bound rule codes in the PR. If an asserted VBE case has no matching
diagnostic implementation, keep it `unbound` or `partially-bound` with a clear
note and create a follow-up issue rather than silently leaving the gap.

## - In final reports, retain the absolute paths of used `tmp_workspaces`, the execution commands, test results, and any unverified items.

## 3. Documentation Retention Policy

### Role Separation Guidelines

- The `tasks/lessons.md` should exclusively serve as a repository for recording recurrence prevention rules - it must not be used for storing design decisions or actual specifications themselves.
- Design judgments and trade-offs should be recorded in the `docs/adr/` directory, while current internal specifications and constraints should be moved to the `docs/specs/` directory.

### Distinction Between ADR Documents and Specification Documentation

- ADR documentation should capture the reasoning behind decisions and document the rationale for choosing one approach over others—information that will be valuable for future implementers facing similar challenges.
- When editing ADR documents, use the `adr-manager` skill.
- Specification documents should record: permanent rules established through review processes, continuous integration testing, and failure resolution; as well as CLI specifications, validation requirements, and compatibility agreements.
- If additional regression tests were added due to specific design considerations where forgetting the rationale could lead to recurrence, document these in the specification documentation.

### Information to Preserve

- The reasoning behind decisions that will be useful for future implementers facing similar issues
- The chosen approach after comparing multiple alternatives
- Permanent rules established through review processes, CI testing, and failure resolution
- CLI specifications, validation requirements, and compatibility agreements
- Documentation of additional regression tests where forgetting the rationale could cause recurrence due to design context

### Information to Discard

- Single-step task notes
- Aborted hypotheses or intermediate notes
- Progress logs that lose value after completion

## - Simple procedure lists without accompanying reasoning justifications

## 4. Core Principles

- **Keep it simple first**: All changes should maintain maximum simplicity with minimal scope impact.
- **Do no harm**: Identify the root cause. Do not resort to quick fixes. Maintain professional engineering standards.
- **Minimize impact**: Only modify necessary components without introducing new bugs.

## 5. Precautions

- Since xlflow operates using both the main binary and the dotnet bridge binary, when performing end-to-end functionality verification, if you install via `go install ./cmd/xlflow`, you will not be able to install the dotnet bridge binary. Always use `task install` for installation.
- When running Go commands on Windows that include `tree-sitter-vba` / CGO, execute them through `scripts/dev/go.ps1` to avoid picking up TDM-GCC. For example: `rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./...`. The commands `task test` / `task install` / `task run` / `task tidy` use this wrapper. If you directly execute `rtk go test` and encounter failures with `runtime/cgo` or `tree-sitter-vba` exiting with status 2, first suspect UCRT64 PATH issues.

## grepai usage

Use `grepai` proactively for semantic code discovery before broad file reads. It
is especially valuable when the task is unfamiliar, spans multiple packages,
describes behavior rather than an exact symbol, or requires architecture and
caller/callee discovery. Do not wait until exact search has already produced a
large candidate set; use semantic search to narrow the search space first.

Recommended flow:

1. Check `rtk grepai status` before searching. If the index is missing or stale,
   start `rtk grepai watch` in a separate terminal and wait for the current
   worktree to be indexed.
2. Use `rtk grepai search "<task intent>"` to find candidate files and concepts.
3. Use `rtk grepai trace callers "<symbol>"` or `rtk grepai trace callees "<symbol>"`
   to identify likely call sites and dependencies.
4. Treat search and trace results as candidates, not ground truth. Confirm the
   important paths with exact search:
   - `rtk rg "<symbol>"`
   - `rtk rg "new <TypeName>"`
   - `rtk rg "<methodName>"`
5. Read only the files confirmed by semantic and exact search, then rerun
   `grepai` after a substantial refactor if the index may have changed.

Use exact `rtk rg` directly when the identifier, config key, diagnostic ID, or
file path is already known; grepai is not a replacement for exact validation.

Branch/index safety:

- Always prefer running `grepai watch` in a separate terminal for multi-step work.
- After `git switch`, validate important grepai hits with `rtk rg` before editing.
- If grepai returns files or symbols that do not exist in the current branch, treat the index as stale and restart `grepai watch`.

Use grepai for:

- semantic code discovery
- finding related implementation files
- locating design/spec documents
- finding likely callers/callees

Use `rtk rg` for:

- exact symbol names
- CLI flags
- error messages
- config keys
- test names

Do not rely on grepai trace alone for complete impact analysis.

## Formatting and pre-commit

- The pre-commit hook must remain non-mutating. It checks only staged files and
  must not run `task fmt` or `goimports -w` automatically.
- Use `pnpm format:check` for the staged-file format check. Use `task fmt`
  explicitly when repository-wide formatting is intentional, then review the
  complete diff before committing.
- Keep Go formatting and import checks read-only in hooks. Do not run overlapping
  formatters in parallel because `gofmt` and `goimports` can both rewrite Go
  files.
- Generated documentation is checked in read-only mode by the lint/docs checks;
  regenerate it explicitly after changing its source registry.

<!-- headroom:rtk-instructions -->

# RTK (Rust Token Killer) - Token-Optimized Commands

When running shell commands, **always prefix with `rtk`**. This reduces context
usage with zero behavior change. If rtk has no filter for a command, it passes
through unchanged — so it is always safe to use.

This project is developed on **Windows**, so prefer PowerShell-compatible
commands and paths.

## Key Commands

```powershell
# Git
rtk git status
rtk git diff
rtk git log --oneline -20

# Files & Search
rtk dir
rtk dir .\src
rtk read .\path\to\file.txt
rtk rg "pattern"
rtk rg "pattern" .\src
rtk find "pattern"
rtk diff .\path\to\file.txt

# Analysis
rtk err <command>
rtk log .\path\to\log.txt
rtk json .\path\to\file.json
rtk summary <command>
rtk deps
rtk env

# GitHub
rtk gh pr view <number>
rtk gh run list
rtk gh issue list
```

<!-- /headroom:rtk-instructions -->
