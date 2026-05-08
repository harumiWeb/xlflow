# ADR-0004: Excel Session Mode and Matching Session Reuse

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

Add explicit session mode for normal agent development, then extend workbook-backed commands to auto-reuse a matching recorded session workbook when that reuse is unambiguous.

`xlflow session start` opens the configured workbook in Excel and records `.xlflow/session.json`. `xlflow push --session`, `xlflow run --session`, and `xlflow save --session` force attachment to that open workbook. When `pull`, `push`, `macros`, `run`, `test`, `trace`, or `save` target the configured workbook and `.xlflow/session.json` still points at the same live workbook, those commands auto-reuse that matching session even if `--session` is omitted. Result payloads and human output surface whether session usage was explicit or auto-reused. `xlflow session status` reports whether the recorded process still exists, whether the workbook is open, and whether the live workbook needs save. `xlflow session stop` saves, closes the workbook, quits Excel, and removes the metadata.

`--session` remains available as an explicit, fail-fast assertion that the matching session workbook must already be running.

The implementation continues to use the existing PowerShell + Excel COM bridge. Session v1 keeps Excel and the workbook alive between commands, but does not introduce a separate Go COM backend.

## Consequences

Positive consequences:

- The default CI and release workflow remains deterministic.
- Agent development loops can avoid repeated Excel startup and workbook open costs.
- Session ownership remains visible in CLI commands and JSON metadata even when auto-reuse occurs.
- The existing script bridge remains the single Excel automation boundary.

Negative consequences:

- Session state can become stale if Excel is killed externally.
- Auto-reuse must be limited to the configured workbook path recorded in `.xlflow/session.json`; anything looser would risk mutating the wrong Excel instance.
- `--session` commands must verify the configured workbook is actually open before mutating it.
- Multiple user-opened Excel instances remain a Windows COM limitation; xlflow v1 validates by workbook path rather than providing a full IPC daemon.
- Unsaved session changes require clear user discipline through `xlflow save` or `xlflow session stop`, plus strong save-required output when live workbook state differs from disk.

## Rationale

- Tests: CLI option tests, script syntax tests, and Excel COM-backed workflow tests.
- Code: `internal/cli`, `internal/excel`, `internal/excel/scripts/session.ps1`, `internal/excel/scripts/push.ps1`, `internal/excel/scripts/run.ps1`.
- Related specs: `docs/specs/cli-contract.md`, `docs/specs/runtime-debugging.md`, `docs/design.md`.

## Supersedes

- None

## Superseded by

- None
