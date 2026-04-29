# xlflow

Agent-ready VBA development framework.

xlflow turns Excel VBA projects into a CLI-first development workflow:

- export VBA modules from `.xlsm`
- edit VBA as normal source files
- import modules back into Excel
- run macros from the CLI
- lint VBA for safer automation
- return deterministic JSON for AI agents

## MVP Commands

```bash
xlflow new
xlflow new Sales
xlflow new Sales.xlsm
xlflow init Book.xlsm
xlflow doctor --json
xlflow pull --json
xlflow push --json
xlflow run Main.Run --json
xlflow lint --json
```

The MVP uses `xlflow.toml` as its project configuration file. Excel automation is Windows-first and uses PowerShell plus Excel COM.
`xlflow new` only accepts `.xlsm` workbook names because it always creates macro-enabled workbook content.

## Verification

Run `task verify` as the obvious local automated verification entry point. It is fast and deterministic, and currently executes non-COM test coverage via `go test ./...`.

```bash
task verify
```

For Excel COM E2E verification, run on a Windows environment with Excel and VBIDE access enabled, then confirm `doctor --json` reports healthy Excel/VBIDE diagnostics.

Use the repo-local `xlflow-tmp-workspace-e2e` skill and run the `tmp_workspaces` flow as the standard real-workbook path. Success criteria are: `new/doctor/pull/lint` all succeed, then module/class/form round-trip checks succeed via `push/run/pull/lint`.
