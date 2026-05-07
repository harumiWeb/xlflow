# xlflow Performance Mode Spec

## Goal

Speed up repeated agent development loops while preserving the current safe defaults for CI and release workflows.

## CLI Contract

- `xlflow push --backup=always|never [--changed-only] [--fast] [--session] [--no-save]`
- `xlflow run [macro] [--direct] [--fast] [--session]`
- `xlflow runner install|remove|status`
- `xlflow session start|status|stop`
- `xlflow save --session`

## Behavior

- Default `push` still backs up, fully replaces non-document components, updates document modules, saves, and closes Excel.
- `push --backup=never` skips backup export.
- `push --changed-only` uses `.xlflow/state/push.json`; unchanged source skips Excel/VBIDE import, changed or missing state falls back to full push.
- `push --fast` expands to `--backup=never --changed-only`.
- `push --no-save` is valid only with `--session`.
- `run --direct` is valid only without `--arg` and without `--trace`; it uses `Excel.Run` directly and returns weaker diagnostics.
- `run --fast` uses direct execution when eligible and otherwise falls back to the normal temporary harness.
- `session start` keeps the configured workbook open and writes `.xlflow/session.json`.
- `--session` commands attach to the already-open workbook by path.
- `session stop` saves, closes, quits Excel, and removes session metadata.

## Outputs

- `push` may include top-level `source.changed`, `source.changed_only`, and `source.state`.
- `run.macro.direct` indicates whether the direct path was used.
- `session` commands may include top-level `session`.
- `runner` commands may include top-level `runner`.

# xlflow Diagnostic Run Spec

## Goal

Add an opt-in diagnostic run mode that converts VBE compile dialogs and VBA runtime failures into structured CLI output for AI agents.

## CLI Contract

- `xlflow run [macro] --diagnostic [--session] [--trace] [--headless | --interactive] [--timeout <duration>]`
- `--diagnostic --direct` is invalid.
- `--diagnostic --fast` is valid, but disables the fast direct path and uses the temporary harness.

## Behavior

- Diagnostic runs keep existing source preflight and GUI-boundary behavior.
- Before harness injection and macro invocation, diagnostic runs execute VBE Compile for the workbook VBA project.
- If VBE raises a modal compile dialog, xlflow reads the dialog text, closes the dialog, reads `ActiveCodePane.GetSelection`, and returns without invoking the macro.
- Runtime failures keep the existing harness-based `macro_failed` result and add `run_diagnostic.kind = "runtime"` when enriched diagnostics are available.
- v1 does not inject line numbers. Runtime line is reported only when VBA `Erl` is non-zero.

## Outputs

- Compile failure returns `error.code = "vba_compile_failed"`, `error.phase = "compile_vba"`, validation exit code `1`, and top-level `run_diagnostic`.
- Compile `run_diagnostic` includes `kind`, `message`, `location`, and `nearby_code` when VBE exposes them.
- Runtime `run_diagnostic` includes `kind = "runtime"` plus existing likely cause, suggestion, nearby source, and trace context fields.

# xlflow Session-Aware Defaults Spec

## Goal

Tighten the normal session workflow so agents can discover the active session automatically, surface save-required state clearly, and inspect richer runtime/build metadata from the CLI.

## CLI Contract

- `xlflow version [--verbose]`
- `xlflow save [--session]`
- `xlflow run [macro] [--session]`
- `xlflow pull|push|macros|test|trace ... [--session]`

## Behavior

- `version --verbose` includes the executable path, Go/build settings, PowerShell script resolution details, and a supported-feature list.
- When `run` is called without a macro argument, it uses `project.entry` from `xlflow.toml`.
- For `pull`, `push`, `macros`, `run`, `test`, `trace`, and `save`, omitting `--session` still reuses the matching live session workbook when `.xlflow/session.json` points at the configured workbook and the session is still open.
- Explicit `--session` and implicit auto-reuse are both preserved in result payloads via workbook session metadata.
- When a session-backed command leaves the live workbook newer than disk, the result includes structured `workbook.needs_save = true` and human output must make the save requirement obvious.
- `session status` includes whether the managed workbook is dirty and whether saving is currently required.
- `push` still saves by default; `--no-save` remains the session-only opt-out.

## Outputs

- Verbose version output may include `version.executable_path`, `version.go_version`, `version.module_path`, `version.build_settings`, `version.scripts`, and `version.features`.
- Workbook-backed commands may include `workbook.session_mode = explicit|auto|managed|none`, plus `workbook.session_requested`, `workbook.auto_session`, and `workbook.needs_save`.
- `session status` may include top-level `session.dirty` and `session.needs_save`.

# Release Artifact Trust Hardening Spec

## Goal

Strengthen trust signals for GitHub Release artifacts without changing the current Windows-first packaging model.

## Release Contract

- GitHub Releases continue to publish `xlflow_windows_x86_64.zip`.
- Releases publish a stable top-level checksum file named `checksums.txt`.
- `checksums.txt` uses SHA256.
- Releases publish SBOM files generated from archive artifacts.
- Release workflow publishes GitHub artifact attestations for release archives, `checksums.txt`, and generated SBOM artifacts.

## User Verification Contract

- Integrity verification: compare the ZIP SHA256 against `checksums.txt`.
- Provenance verification: run `gh attestation verify <zip> --repo harumiWeb/xlflow`.
- Non-claim: neither verification step is documented as proof of Windows publisher identity or Authenticode signing.
