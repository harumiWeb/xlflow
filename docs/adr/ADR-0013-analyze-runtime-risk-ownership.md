# ADR-0013: Analyze Owns Semantic VBA Runtime-Risk Checks

## Status

`accepted`

## Background

xlflow has two source-only feedback commands:

- `lint` reports syntax-safety, style, parser recovery, and local code-shape issues.
- `analyze` reports likely runtime failures and source patterns that can explain or prevent Excel/VBA runtime errors.

The AST-backed lint work temporarily included several runtime-risk checks, including unqualified Excel object access, error-handler fallthrough, Application state restore, and `Range.Find` results used without `Nothing` guards. Those checks need declaration, procedure, and flow context, and they overlap with analyzer diagnostics used by source preflight and runtime failure suggestions.

The relevant lint/analyze rule set is still under `## Unreleased`, so xlflow does not need compatibility aliases or a migration period for the temporary lint placement.

## Decision

`xlflow analyze` owns semantic VBA runtime-risk checks that require declaration, procedure, call, assignment, label, member-access, or return-path context.

`xlflow lint` remains focused on syntax-safety, parser recovery, local code-shape, and automation-hostile GUI boundaries. Runtime-risk checks previously staged in lint before release are moved to analyze without compatibility aliases.

The analyzer uses `tree-sitter-vba` through `internal/vba/ast` and keeps the existing analysis finding contract: `code`, `severity`, `file`, `module`, `procedure`, `line`, `message`, `reason`, `suggestion`, and `nearby_code`, with `column` included only when reliable.

Stable analyzer codes include existing `VBA101` through `VBA106` plus runtime-risk `VBA201` through `VBA211`.

## Consequences

Positive consequences:

- `check` output separates local lint feedback from runtime-risk analysis more clearly.
- Analyzer findings can be reused by source preflight and runtime failure suggestions without duplicating lint issues.
- AST-backed analysis can ignore comments and strings while using procedure and declaration context.

Negative consequences:

- The analyzer becomes more complex and needs focused tests for conservative false-positive behavior.
- Documentation and generated config must describe both `[lint]` and `[analyze]` rule families.
- False-positive-prone analyzer rules need opt-in defaults until their behavior is proven.

## Rationale

- Specs: `docs/specs/cli-contract.md`, `docs/specs/runtime-debugging.md`.
- Docs: `vitepress/commands/lint.md`, `vitepress/commands/analyze.md`, `vitepress/reference/config-file.md`.
- Tests: analyzer regression tests for `VBA101` through `VBA106` and runtime-risk tests for `VBA201` through `VBA211`.
- Code: `internal/analyze`, `internal/lint`, `internal/config`.

## Supersedes

- None

## Superseded by

- None
