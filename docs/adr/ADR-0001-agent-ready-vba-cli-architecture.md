# ADR-0001: Agent-ready VBA CLI Architecture

## Status

`accepted`

## Background

xlflow is intended to make Excel VBA projects editable, testable, and repairable by AI agents. The core need is deterministic CLI behavior, stable machine-readable output, and a bridge to Excel/VBIDE without hiding the Windows and COM dependency.

Alternative approaches considered:

- A pure Go implementation for all workbook manipulation.
- A thick cross-platform adapter abstraction from the start.
- A Windows-first Go CLI with PowerShell scripts handling Excel COM operations.
- Treating Excel/VBA GUI operations as invisible automation targets instead of explicit boundaries.

The MVP needs useful behavior quickly while keeping the public contracts stable.

## Decision

Adopt a Go CLI using Cobra for command routing, `xlflow.toml` for project configuration, stable JSON envelopes for AI consumption, and PowerShell scripts as the Excel COM bridge.

The MVP is Windows-first. Excel-dependent commands are expected to run on Windows with Excel installed and suitable VBIDE trust settings. Non-Excel commands such as `init` and `lint` remain testable without Excel.

`push` creates a timestamped backup under `.xlflow/backups/` before replacing VBA components.

VBA source-controlled files use UTF-8 without BOM. The Excel COM bridge treats VBIDE import and export text files as CP932 and converts at the `pull`/`push` boundary. Binary userform companions such as `.frx` are copied without text conversion.

GUI-dependent VBA calls are treated as explicit automation boundaries. xlflow detects patterns such as file pickers, modal dialogs, UserForms, message pumps, and external process launches in source code. Headless runs reject these boundaries before Excel starts. Interactive runs make the human participation explicit by showing Excel and allowing alerts. xlflow does not attempt to silently automate arbitrary Excel dialogs or UserForms.

VBE compile-error dialogs are a narrow exception because they are editor diagnostics, not workbook business UI. `run --diagnostic` may execute VBE Compile, read and close the VBE modal compile dialog for the owned Excel process, and return structured compile diagnostics. This does not extend to user-code prompts such as `MsgBox`, file pickers, or UserForms.

## Consequences

Positive consequences:

- AI agents can rely on stable JSON and exit-code contracts.
- Excel automation stays isolated in scripts that can be inspected independently.
- The Go code remains focused on CLI, configuration, validation, and result handling.
- GUI-dependent workflows remain usable with a human in the loop while keeping unattended runs predictable.

Negative consequences:

- Excel operations are not portable beyond Windows in the MVP.
- PowerShell and COM behavior must be tested separately from pure Go logic.
- Japanese and other non-ASCII VBA source depends on the bridge conversion staying explicit at import and export boundaries.
- Static GUI boundary detection can produce conservative findings and does not prove that a specific macro path will show a dialog.
- VBE compile dialog handling is Windows-specific and depends on Win32 window enumeration in addition to Excel/VBIDE COM.
- Future non-Windows backends may require a new ADR before expanding the adapter boundary.

## Rationale

- Tests: Go unit tests for configuration, output, project scaffolding, and lint behavior; PowerShell syntax checks for bridge scripts.
- Code: `cmd/xlflow`, `internal/cli`, `internal/excel`, `internal/excel/scripts/*.ps1`.
- Related specs: `docs/specs/cli-contract.md`, `docs/specs/runtime-debugging.md`, `docs/design.md`.

## Supersedes

- None

## Superseded by

- None
