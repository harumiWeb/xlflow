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
