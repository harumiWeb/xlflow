# xlflow Architecture

## Core Shape

- `cmd/xlflow` wires the CLI entrypoint.
- `internal/cli` owns user-facing command parsing and argument validation.
- `internal/excel` owns workbook automation, PowerShell bridge invocation, and Excel COM behavior.
- `internal/excel/scripts` contains shared PowerShell bridge code that must stay compatible with Windows PowerShell 5.1.
- `internal/project` owns project scaffolding and source layout decisions.
- `internal/output` owns human and JSON-facing rendering behavior.
- `docs/specs` stores durable CLI, validation, and compatibility contracts.
- `docs/adr` stores architecture decisions and tradeoffs.
- `vitepress` and README files store user-facing documentation.

## Architectural Principles

- Workbook-first architecture: commands must respect the configured workbook and active session state.
- Deterministic export pipeline: repeated pull/export/snapshot operations should produce stable source artifacts.
- Stable VBA generation: generated code should be explicit, workbook-qualified where needed, and cleanup-safe.
- Spec-driven form runtime: UserForm specs, designer snapshots, `.frm`, `.frx`, and sidecar code must stay coherent.
- Boundary resolution: CLI/config/env/path choices should be normalized at command boundaries before core execution.
- Preflight before Excel: predictable source and config errors should be reported before opening Excel or invoking VBIDE.

## Workbook Command Boundaries

- Inspection opens and execution opens have different macro-safety requirements.
- `run`, `test`, and `session start` must keep target macros executable.
- Non-executing inspection paths can force-disable automation macros.
- `--session` commands must reattach to the exact Excel instance recorded by `session start`.
- Trace, inspect, edit, and save flows must not silently operate on a hidden second workbook copy when a matching live session exists.

## UserForm Source Model

- UserForm designer state and code-behind are separate source-controlled artifacts.
- In sidecar mode, `src/forms/code/*.bas` is authoritative for code-behind.
- Tracked `.frm` files in sidecar mode are generated artifacts for diagnostics and compatibility.
- Spec filename, `form.name`, `.frm` basename, and `.frm` `Attribute VB_Name` must agree before import/build.
- Existing projects must not silently switch authoritative UserForm source mode.

## Documentation Boundaries

- Use `tasks/todo.md` for session progress and temporary verification notes.
- Use `tasks/feature_spec.md` for working specs while implementing.
- Promote long-lived behavior, CLI contracts, and validation rules to `docs/specs/`.
- Use ADRs for tradeoffs or architecture decisions that future maintainers should understand.
