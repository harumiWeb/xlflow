# ADR-0004: Explicit Excel Session Mode

## Status

`accepted`

## Background

`xlflow push` and `xlflow run` historically open Excel, open the configured workbook, perform one operation, then close Excel. This is safe and deterministic, but it is expensive for agent-driven loops such as `push -> run -> edit -> push -> run`.

Alternatives considered:

- Keep every command isolated and only add small flags such as `--backup=never` and `--direct`.
- Automatically reuse any already-running Excel instance for normal `push` and `run`.
- Add an explicit session mode that keeps the configured workbook open and requires callers to opt in with `--session`.
- Replace the PowerShell bridge with direct Go COM automation.

## Decision

Add explicit session mode while keeping the default command behavior isolated.

`xlflow session start` opens the configured workbook in Excel and records `.xlflow/session.json`. `xlflow push --session`, `xlflow run --session`, and `xlflow save --session` attach to that open workbook. `xlflow session status` reports whether the recorded process still exists and whether the workbook is open. `xlflow session stop` saves, closes the workbook, quits Excel, and removes the metadata.

Normal `push` and `run` do not automatically use a session. The user or agent must opt in with `--session`.

The implementation continues to use the existing PowerShell + Excel COM bridge. Session v1 keeps Excel and the workbook alive between commands, but does not introduce a separate Go COM backend.

## Consequences

Positive consequences:

- The default CI and release workflow remains deterministic.
- Agent development loops can avoid repeated Excel startup and workbook open costs.
- Session ownership is explicit in CLI commands and JSON metadata.
- The existing script bridge remains the single Excel automation boundary.

Negative consequences:

- Session state can become stale if Excel is killed externally.
- `--session` commands must verify the configured workbook is actually open before mutating it.
- Multiple user-opened Excel instances remain a Windows COM limitation; xlflow v1 validates by workbook path rather than providing a full IPC daemon.
- Unsaved session changes require clear user discipline through `xlflow save --session` or `xlflow session stop`.

## Rationale

- Tests: CLI option tests, script syntax tests, and Excel COM-backed workflow tests.
- Code: `internal/cli`, `internal/excel`, `scripts/session.ps1`, `scripts/push.ps1`, `scripts/run.ps1`.
- Related specs: `docs/specs/cli-contract.md`, `docs/design.md`.

## Supersedes

- None

## Superseded by

- None
