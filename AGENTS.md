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

## - In final reports, retain the absolute paths of used `tmp_workspaces`, the execution commands, test results, and any unverified items.

## 3. Documentation Retention Policy

### Role Separation Guidelines

- The `tasks/todo.md` file may temporarily contain not only session-level progress tracking but also temporary records of verification outcomes, unresolved issues, and summary explanations of decision rationales.
- The `tasks/feature_spec.md` may be used as a working specification document prior to implementation, but should not be discarded if it contains future reference specifications, constraints, or test conditions.
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

Use `grepai` for semantic code discovery before broad file reads.

Recommended flow:

1. Use `grepai search "<task intent>"` to find candidate files.
2. Use `grepai trace callers "<symbol>"` or `grepai trace callees "<symbol>"` to identify likely call sites.
3. Treat trace results as candidates, not ground truth.
4. Verify important symbols and call sites with exact search:
   - `rtk rg "<symbol>"`
   - `rtk rg "new <TypeName>"`
   - `rtk rg "<methodName>"`
5. Read only the files confirmed by grepai + exact search.

Branch/index safety:

- Prefer running `grepai watch` in a separate terminal.
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
