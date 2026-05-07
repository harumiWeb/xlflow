# Changelog

All notable changes to xlflow will be documented in this file.

## v0.5.0

- Added richer sample VBA projects, including the `world-news` NewsAPI example and the `stock-price` dashboard example, plus accompanying screenshots and README updates.
- Improved runtime error handling and diagnostics so CLI runs surface failures more clearly across the Go and PowerShell execution bridge.
- Refined release documentation and sample project metadata with formatting fixes and README polish, including Japanese README badge updates.

## v0.4.0

- Added `xlflow inspect` with workbook, sheet, range, used-range, and cell inspection for saved workbook snapshots.
- Added inspect-specific formatting and range limits so agents can read workbook structure and output without opening Excel.
- Updated the bundled xlflow agent skill and CLI contract docs to teach snapshot-first inspect workflows.

## v0.3.0

- Added automatic reuse of a matching live xlflow session workbook for workbook-backed commands when `--session` is omitted.
- Added structured save-state reporting so `push`, `run`, `session status`, and related commands can surface when a live session workbook differs from disk and needs `xlflow save`.
- Improved `run` with compile-first diagnostic mode, clearer direct-run restrictions, and fallback to `project.entry` when no macro argument is provided.
- Expanded trace lifecycle handling with enable/disable/status/clean flows, temporary trace injection, and session-aware workbook reuse.
- Added a verbose `version` command that reports build metadata, script resolution, supported features, and executable details.
- Added update-checking and refreshed version/welcome messaging.
- Updated bundled PowerShell scripts, agent skill guidance, and JSON envelopes to match the new session-aware behavior.

## v0.2.0

- Bundled the PowerShell scripts used by xlflow for Excel session management, testing, tracing, and UI button manipulation.
- Added the initial session-aware command surface for opening, reusing, saving, and stopping Excel workbooks.
- Added trace, run, pull, push, test, and UI button workflows built on the bundled PowerShell bridge.
