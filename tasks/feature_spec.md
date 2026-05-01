# Runtime Diagnostics Spec

## Goal

Make xlflow failures easier to repair by combining lightweight VBA source analysis, aggregate preflight checks, lower-friction trace commands, and enriched `run` failure diagnostics.

## CLI Contract

- `xlflow analyze [--json]`
- `xlflow check [--json] [--keepalive] [--keepalive-interval <duration>]`
- `xlflow trace enable [workbook] [--keepalive] [--keepalive-interval <duration>]`
- `xlflow trace disable [workbook] [--force] [--keepalive] [--keepalive-interval <duration>]`
- `xlflow trace status [workbook] [--keepalive] [--keepalive-interval <duration>]`
- `xlflow trace clean [--keepalive] [--keepalive-interval <duration>]`
- `xlflow trace inject [workbook]` remains a compatibility alias for `trace enable`.

## Output Contract

- `analyze` returns top-level `analysis`.
- `check` returns top-level `check`, plus `issues`, `analysis`, and doctor diagnostics where available.
- Failed `run` results may include top-level `run_diagnostic`.
- Trace commands return lifecycle/status/clean metadata under top-level `trace`.

## Analyzer Rules

- `VBA101`: object variable assignment likely missing `Set`.
- `VBA102`: object-returning function assignment likely missing `Set`.
- `VBA103`: object-returning function body likely missing `Set <FunctionName> = ...`.

## Failure Contract

- Analyzer findings exit `1`.
- `check` exits `1` for lint/analyze findings and `3` for doctor/environment failure.
- `trace disable` refuses to delete a modified `XlflowTrace.bas` unless `--force` is set.
