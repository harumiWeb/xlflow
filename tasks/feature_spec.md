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
