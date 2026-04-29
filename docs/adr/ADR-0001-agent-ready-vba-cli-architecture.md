# ADR-0001: Agent-ready VBA CLI Architecture

## Status

`accepted`

## Background

xlflow is intended to make Excel VBA projects editable, testable, and repairable by AI agents. The core need is deterministic CLI behavior, stable machine-readable output, and a bridge to Excel/VBIDE without hiding the Windows and COM dependency.

Alternative approaches considered:

- A pure Go implementation for all workbook manipulation.
- A thick cross-platform adapter abstraction from the start.
- A Windows-first Go CLI with PowerShell scripts handling Excel COM operations.

The MVP needs useful behavior quickly while keeping the public contracts stable.

## Decision

Adopt a Go CLI using Cobra for command routing, `xlflow.toml` for project configuration, stable JSON envelopes for AI consumption, and PowerShell scripts as the Excel COM bridge.

The MVP is Windows-first. Excel-dependent commands are expected to run on Windows with Excel installed and suitable VBIDE trust settings. Non-Excel commands such as `init` and `lint` remain testable without Excel.

`push` creates a timestamped backup under `.xlflow/backups/` before replacing VBA components.

## Consequences

Positive consequences:

- AI agents can rely on stable JSON and exit-code contracts.
- Excel automation stays isolated in scripts that can be inspected independently.
- The Go code remains focused on CLI, configuration, validation, and result handling.

Negative consequences:

- Excel operations are not portable beyond Windows in the MVP.
- PowerShell and COM behavior must be tested separately from pure Go logic.
- Future non-Windows backends may require a new ADR before expanding the adapter boundary.

## Rationale

- Tests: Go unit tests for configuration, output, project scaffolding, and lint behavior; PowerShell syntax checks for bridge scripts.
- Code: `cmd/xlflow`, `internal/cli`, `internal/excel`, `scripts/*.ps1`.
- Related specs: `docs/specs/cli-contract.md`, `docs/design.md`.

## Supersedes

- None

## Superseded by

- None
