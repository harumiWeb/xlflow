# xlflow Runtime Debugging Hardening Spec

## Skill provider target correction

Copilot must not be offered as a provider-specific skill target. GitHub Copilot reads repository instructions from `.agents`, so users should choose `agents` for shared Copilot-compatible instructions or pass an explicit `--target` when they need a custom location. The supported `--agent` values for bundled skill installation are `agents`, `codex`, `claude`, `cursor`, and `gemini`.

## Goal

Make xlflow easier for AI agents to use when debugging workbook macros from the CLI. The main failure mode to address is that source-controlled VBA and workbook VBA can drift, and runtime failures currently require too much implicit knowledge to interpret.

## Required Behavior

### Trace injection persistence

- `xlflow trace inject` must keep the configured project source tree and workbook in sync.
- When a project configuration is available and no explicit workbook argument is provided, `trace inject` must:
  - inject or replace the `XlflowTrace` standard module in the configured workbook,
  - write the same module source to `<src.modules>/XlflowTrace.bas`,
  - encode the source-controlled `.bas` file as UTF-8 without BOM,
  - report the generated source path in JSON output.
- When an explicit workbook argument is provided and no project configuration is available, `trace inject <workbook>` may operate on the workbook only.
- Re-running `trace inject` must be idempotent. It may replace `XlflowTrace.bas` with the bundled trace module source.
- After `trace inject` in a configured project, a subsequent `xlflow push` must preserve the trace module instead of deleting it.

### Run failure diagnostics

- `xlflow run` failures must identify the phase that failed.
- JSON failures should include a stable phase value such as:
  - `open_workbook`
  - `prepare_vbide`
  - `inject_harness`
  - `invoke_macro`
  - `save_result`
  - `read_trace`
- Macro invocation failures caused by a missing or invalid target macro should be distinguishable from failures raised by user code when Excel exposes enough information to make that distinction.
- Plain-text failures should remain concise but should include enough context for an AI agent to choose the next command or source file to inspect.

### Macro entrypoint discovery

- xlflow must provide a machine-readable way to discover runnable macro entrypoints before calling `run`.
- The discovery result should include at least:
  - module name,
  - procedure name,
  - fully qualified macro name,
  - procedure kind when available,
  - argument count or argument signature when available.
- The command must avoid executing user code.
- Discovery should read the configured workbook by default and support JSON output.
- The bundled AI agent skill must instruct agents to use macro discovery before guessing a `run` target when the entrypoint is unclear.

### Automation-hostile VBA detection

- `xlflow lint` should detect VBA patterns that block unattended CLI execution.
- Initial patterns:
  - `Application.GetOpenFilename`
  - `Application.FileDialog`
  - `InputBox`
  - modal `MsgBox` usage
- Findings should explain that CLI-oriented macros should prefer explicit arguments, environment variables, configuration cells, or deterministic paths.
- The first implementation may treat these as lint failures to match the existing lint model, but the rule text must make the automation concern clear.
- The bundled AI agent skill must tell agents to avoid designing macros around UI prompts when CLI execution is expected.

### Empty trace guidance

- `xlflow run --trace` must preserve the existing behavior of returning trace events written before a failure.
- If a traced run fails with zero events, output should make it clear that execution may have failed before reaching user trace calls.
- The bundled AI agent skill should tell agents to add trace logs at procedure entry, key branches, external file access, and error handlers.
- The bundled AI agent skill should tell agents how to interpret `macro_not_found`, run failure phases, and empty trace results.

## Verification

- Unit tests must cover source-file creation for `trace inject` without requiring Excel where practical.
- PowerShell helper tests must cover the generated trace module source used for both workbook injection and source persistence.
- CLI or bridge tests must cover the JSON fields added for trace source persistence and run failure phases.
- Lint tests must cover the new automation-hostile VBA patterns.
- Documentation must be updated in `docs/specs/cli-contract.md`, `README.md` where user-facing command behavior changes, and the bundled `xlflow` skill where AI workflow guidance changes.
