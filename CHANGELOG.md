# Changelog

All notable changes to xlflow will be documented in this file.

## Unreleased

- Added native `.NET` bridge support for `xlflow macros --bridge dotnet --json` and `xlflow run Module1.Main --bridge dotnet --json`, enabling macro discovery and execution through the .NET Excel bridge without PowerShell. Supports typed arguments including finite invariant-culture `double` values, fully qualified macro names, save/no-save/save-as, timeout, session attachment, and structured error handling for `macro_failed`, `macro_not_found`, and `macro_disabled`. Auto mode keeps the existing PowerShell behavior for macros/run; use `--bridge dotnet` explicitly to route through the .NET bridge.
- Added a reusable `.NET` Excel/VBE dialog watcher that captures runtime, compile, MsgBox, InputBox, and FileDialog snapshots with Win32/UI Automation identity metadata. Runtime error dialogs are suppressed without requiring Excel focus, and unattended runs prefer End over Debug to avoid leaving VBE in break mode.
- Added native `.NET` bridge support for `xlflow pull --bridge dotnet --json` and `xlflow push --bridge dotnet --json`, enabling VBA component export/import through the .NET Excel bridge without PowerShell. Auto mode keeps the existing PowerShell behavior for pull/push; use `--bridge dotnet` explicitly to route through the .NET bridge.
- Added native `.NET` bridge support for runner-backed `xlflow inspect workbook|sheets|range --session --bridge dotnet --json` and `xlflow process list|cleanup --bridge dotnet --json`, including `--bridge auto` fallback from unsupported/runtime/protocol `.NET` failures back to PowerShell for supported commands.
- Added native `.NET` `xlflow doctor --bridge dotnet --json` diagnostics for runtime and Excel COM probing, plus documentation clarifying that top-level `bridge` metadata remains provider-specific between PowerShell and `.NET` bridges.
- Added structured COM error fields (`h_result`, `details`) to `xlflow doctor --bridge dotnet --json` error output. COM activation failures now include the HRESULT hex code and exception details alongside the error message.
- Added an Excel bridge provider abstraction in Go, moved PowerShell invocation behind `PowerShellProvider`, and added bridge selection via persistent `--bridge`, `XLFLOW_EXCEL_BRIDGE`, and `[excel].bridge` while keeping `auto` on the existing PowerShell behavior for now.
- Added `xlflow fmt` as a conservative, non-destructive VBA source formatter for `.bas` and `.cls` files. Supports `--write`, `--check`, `--diff`, `--json`, and `--stdin` modes. The formatter uses 4-space indentation, strips trailing whitespace, normalizes blank lines, preserves class module metadata, and is idempotent. Typical workflow: `fmt -> lint -> push -> run/test`.
- Refined the interactive `xlflow new` / `init` welcome screen with a new `Welcome to` heading, a command reference URL, and softer muted version/info text below the ASCII logo.
- Hardened the bundled TAKT orchestra, PR review, and issue bug workflows with explicit verification, audit-triage, and release-gate handling, broader loop monitoring around remediation and final audit, and clearer guidance to treat allowed untracked files and auto-staging state as non-blocking.
- Added `xlflow process list` to enumerate all local Excel processes with PID and open-workbook status.
- Added `xlflow process cleanup <pid>`, `xlflow process cleanup --auto`, and `xlflow process cleanup --all [--yes]` for safe and forceful Excel process termination. `--auto` targets only workbook-free processes; `--all` is a destructive force-stop of all local Excel instances with mandatory interactive confirmation or `--yes`.
- Fixed `XlflowDebug.bas` helper module to stop forwarding `Log`'s `ParamArray` into a secondary helper procedure, preventing VBA compile/runtime failures such as "Sub または Function が定義されていません" and "ParamArray の使い方が適切ではありません" in some hosts.
- Fixed `xlflow run --diagnostic` compile watcher to return structured `vba_compile_failed` errors when the VBE compile command control cannot be found, instead of silently reclassifying the failure as `vbide_access_denied`.
- Improved runtime dialog capture for `xlflow run --diagnostic` so break-mode inspection prefers user-code lines over temporary `XlflowRun_*` helpers, and the runtime macro runner now executes in a disposable child PowerShell process so break-mode resets do not leave the parent CLI hung.
- Fixed `xlflow run --diagnostic` VBE compile preflight to locate `Compile VBAProject` from the VBE menu bar (`Id = 578`) instead of assuming the Debug toolbar contains it, and to treat a disabled compile command as "already compiled" rather than a hard failure.
- Fixed `xlflow ui button add` so it auto-reuses a matching live session workbook when `.xlflow/session.json` points at the configured workbook, preventing the Excel SaveAs dialog that previously appeared when a session was active.
- Extended `ui button add`, `ui button list`, and `ui button remove` to use the shared session-aware workbook open helper and explicit save/release cleanup, matching the behavior of `push`, `pull`, `run`, `trace`, and other workbook-backed commands.
- Added `xlflow status` and `xlflow status --json` as a read-only project-level command that shows project, source, workbook, and session state in one output. Source freshness is a heuristic based on file mtimes; the command does not modify workbook files, source files, or `.xlflow/state`. `workbook_saved` is now derived from `save_required` instead of `dirty` to avoid contradictory results when the session probe reports `save_required=true` but `dirty` is unknown; baseline `session` payload now always includes `running`, `workbook_open`, and `metadata` for schema stability.
- Added `xlflow init --with-module` so imported projects can immediately receive bundled runtime helper modules and sync them back into the copied workbook.
- Added `xlflow module install [--push]` so existing xlflow projects can install bundled helper modules on demand without rerunning `new`.
- Removed `--keepalive` / `--keepalive-interval` from Excel COM-backed commands and the final `XLFLOW_DONE` marker; interactive stderr now uses spinner progress where available, while non-interactive runs fall back to line-oriented progress and streamed UI/debug stderr output suppresses separate progress frames.
- Added XlflowUI module with MsgBox and InputBox wrappers to handle user prompts.
- Extended XlflowUI with headless-safe file dialog wrappers for `Application.GetOpenFilename`, `Application.GetSaveAsFilename`, open `Application.FileDialog`, and folder picker flows, plus repeated `--filedialog <kind>:<dialog-id>=<value>` CLI responses for `run` and `test`.
- Added `--ui-stream` for `xlflow run` and `xlflow test`, streaming resolved headless `XlflowUI` dialog events to stderr in real time while preserving JSON stdout and returning final `ui.events` payloads plus human-readable `UI` summaries.
- Added scaffolded `XlflowDebug` helper support so explicit `XlflowDebug.Log` calls stream to stderr and final top-level `debug` payloads during `xlflow run` and `xlflow test` without a separate CLI flag, including direct and fast run paths.
- Updated run.ps1 and test.ps1 to accept MsgBoxResponsesJSON and InputResponsesJSON parameters.
- Added explanatory comments to scaffolded `XlflowRuntime.bas`, `XlflowUI.bas`, and `XlflowAssert.bas` so workbook authors can adopt the helper modules more easily.
- Added explicit live-session inspect mode for `inspect workbook`, `inspect sheets`, `inspect range`, `inspect used-range`, and `inspect cell` via `--session`, plus explicit `live_session` target metadata and saved-file warnings that point callers to live-session inspect when disk may be stale.
- Added runtime-aware execution mode injection for `run` and `test`, plus the scaffolded `XlflowRuntime` VBA helper for branching between interactive, headless, agent, CI, and test execution contexts.
- Enhanced `xlflow macros --json` output with `component_type`, `visibility`, `has_parameters`, `runnable`, `reason_not_runnable`, and `run_command` fields per macro so AI agents and users can choose the correct entrypoint without guessing.
- Added `default_entry` and `suggestions` fields to `xlflow macros --json` output, surfaced from `project.entry` in `xlflow.toml` and resolved against discovered runnable macros.
- Added `--runnable` flag to `xlflow macros` to filter the output to only directly runnable procedures.

## v0.9.0

- Added winget release publishing so GoReleaser can generate the `HarumiWeb.Xlflow` manifest and push it to the `harumiWeb/winget-pkgs` fork for upstream submission.
- Updated `xlflow new` to bootstrap the workbook/source sync automatically by pushing scaffolded VBA into the new workbook before the command reports success, and added placeholder `src/workbook/ThisWorkbook.bas` / `Sheet1.bas` files with `Option Explicit` for new projects.
- Updated `xlflow init` to bootstrap source sync automatically by pulling VBA from the copied workbook into `src/`.
- Added first-class workbook rollback support with `xlflow backup list` and `xlflow rollback`, including metadata-backed workbook-file backups under `.xlflow/backups/<backup-id>/`, automatic safety backups before restore, and session-aware guards that refuse rollback while the target workbook is open in an active xlflow session.
- Changed default `push` backups from component-export snapshots to rollback-capable workbook snapshots, and updated the CLI/docs surface, JSON output, and VitePress command/concept pages to reflect the new backup and recovery workflow.

## v0.8.1

- Fixed `xlflow inspect form <name> --designer --session` so normal designer inspection no longer takes the strict temporary-workbook path, reducing the sample `space-invader` session inspection from about one minute to a few seconds.
- Corrected PowerShell boolean parsing and case-insensitive variable handling around the `StrictDesigner` flag, preventing `"False"` string values from being treated as truthy.
- Hardened UserForm runtime cleanup guards in `inspect form` and `form export-image` so null runtime workbook state does not trigger unnecessary Excel COM cleanup and finalizer waits.

## v0.8.0

- Completed the UserForm feature set for issue #25 across phase 1 through phase 7, including explicit UserForm warnings in core workbook flows, `xlflow list forms`, `inspect form` for designer/runtime/both, `form snapshot`, and experimental runtime image export.
- Hardened `form export-image` for real Excel GUI behavior by repairing generic runtime captions from designer state, choosing the correct monitor-relative work area instead of forcing the primary screen, using DPI-aware capture sizing, and trimming capture artifacts so the exported PNG matches the visible UserForm more faithfully.
- Corrected UserForm build round-tripping so snapshot-derived width and height no longer grow on each `form build` cycle, preserving the persisted Designer dimensions from `snapshot` output.
- Updated the bundled docs, CLI contract, and agent guidance to reflect the UserForm discovery, inspection, snapshot, export, and warning workflow, including the experimental status of runtime image export.
- Strengthened PowerShell script coverage with behavior-oriented tests for the UserForm build and export helpers, replacing narrow string-presence checks where practical.

## v0.7.0

- Added `xlflow edit cell`, `edit range`, `edit rows`, and `edit columns` as minimal workbook-mutation helpers for AI-agent testing and visual tuning in a live Excel session.
- Added session-only workbook edit behavior for the new `edit` commands, including `--events keep|on|off` support for cell value and formula changes so `Worksheet_Change` flows can be exercised without generating temporary VBA.
- Commands now display explicit workbook state, including whether reading from saved file or live Excel session
- Added warnings when live session workbooks contain unsaved changes
- Extended workbook-backed JSON and human output with explicit `target` / `session` metadata across session-aware commands, plus top-level `edit` payloads for workbook mutation summaries.
- Updated the CLI contract, README files, ADR session policy note, and bundled xlflow skill guidance to cover the new edit workflow and session-state visibility.

## v0.6.0

- Added `xlflow export-image` to export worksheet ranges as PNG images for visual verification, including session-aware targeting, structured `target` / `output` metadata, and reliability fixes so hidden-workbook captures do not produce blank images or leak Excel processes.
- Added `--include-style` flag to `inspect range` and `inspect used-range` commands to display worksheet style metadata including cell fills, borders, merged cells, row heights, and column widths.
- Added Rubberduck-compatible folder-aware VBA sync so `xlflow pull` and `push` can round-trip nested source trees via `@Folder(...)`, recursive source discovery, duplicate module-name preflight, and nested `.frm`/`.frx` companion handling.
- Added `[vba]` configuration defaults for folder sync control, wired the settings through the Go/PowerShell bridge, and documented the new contract in the CLI spec, READMEs, and bundled xlflow skill.
- Fixed folder-sync path handling to stay compatible with Windows PowerShell 5.1 and hardened `pull` so it does not clear the existing exported source tree before the workbook opens successfully.
- Added `--no-update-check` and `XLFLOW_NO_UPDATE_CHECK=1` so interactive `new` and `init` can skip the GitHub Release lookup used by the scaffold welcome banner.
- Hardened GitHub Release packaging with stable `checksums.txt` SHA256 output and archive SBOM generation via GoReleaser.
- Extended the release workflow to install Syft and publish GitHub artifact attestations for release archives, checksums, and SBOM artifacts.
- Documented Windows-side release verification in both READMEs, including SHA256 checks, `gh attestation verify`, and the current non-goal of Authenticode signing.

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
