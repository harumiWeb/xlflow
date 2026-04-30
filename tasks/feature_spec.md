# xlflow Trace Harness Spec

## Goal

Add trace logging support so agents and humans can collect simple VBA execution events during `xlflow run`.

## Behavior

- `xlflow trace inject [workbook]` injects or replaces a standard VBA module named `XlflowTrace`.
- When workbook is omitted, `trace inject` uses `[excel].path` from `xlflow.toml`; when workbook is explicit, it can run without project configuration.
- The injected module exposes `XlflowSetTraceFile path` for the run harness and `XlflowLog message` for user VBA code.
- `xlflow run [macro] --trace` creates a fresh temp trace log under `%TEMP%\xlflow`, configures `XlflowTrace`, runs the macro, then reads trace events back into CLI output.
- `run --trace` fails with `trace_not_injected` when the target workbook does not contain `XlflowTrace`.
- Macro failures still include any trace events written before the failure.

## Interfaces

- CLI: `xlflow [--json] trace inject [workbook]`
- CLI: `xlflow [--json] run [macro] [--input <workbook>] [--arg <type:value>]... [--save | --save-as <path>] [--trace]`
- VBA: `Call XlflowLog("message")`
- JSON: top-level `trace` field containing `enabled`, `path`, `events`, and optional `read_error`.
- Trace events contain `timestamp`, `message`, and `raw`.

## Verification

- Fast gate: `go test ./...` and `task verify`.
- Script gate: PowerShell parse tests include `trace.ps1`; helper tests cover trace module generation, trace harness setup, and event parsing.
- Integration gate: create a temporary workbook, run `trace inject`, add `XlflowLog` calls, run `xlflow run --trace --json`, and confirm `trace.events` contains the expected messages.
