# xlflow

Agent-ready VBA development framework.

xlflow turns Excel VBA projects into a CLI-first development workflow:

- export VBA modules from `.xlsm`
- edit VBA as normal source files
- import modules back into Excel
- run macros from the CLI
- run VBA tests from the CLI
- compare workbook state and exported VBA source
- lint VBA for safer automation
- return deterministic JSON for AI agents

## MVP Commands

```bash
xlflow new
xlflow new --with-skill --agent codex
xlflow new Sales
xlflow new Sales.xlsm
xlflow init Book.xlsm
xlflow init Book.xlsm --with-skill --agent claude
xlflow doctor --json
xlflow pull --json
xlflow push --json
xlflow trace inject
xlflow run Main.Run --json
xlflow run Main.Run --trace --json
xlflow run Report.Generate --arg string:fixtures\sample.xlsx --arg int:3 --arg bool:true --save --json
xlflow run Report.Generate --input build\Book.xlsm --arg string:fixtures\sample.xlsx --save-as build\Result.xlsm --json
xlflow test --json
xlflow test --filter TestCreateReport --json
xlflow diff before.xlsm after.xlsm --json
xlflow diff before.xlsm after.xlsm --vba-before before-src --vba-after after-src --json
xlflow lint --json
xlflow skill install --agent codex
xlflow skill install --target .agents/skills
```

The MVP uses `xlflow.toml` as its project configuration file. Excel automation is Windows-first and uses PowerShell plus Excel COM.
`xlflow new` only accepts `.xlsm` workbook names because it always creates macro-enabled workbook content.

`xlflow new/init --with-skill` installs the bundled `xlflow` AI agent skill during setup. The same skill can be installed later with `xlflow skill install`. Supported provider targets are `agents`, `codex`, `claude`, `cursor`, `gemini`, and `copilot`; when no provider is specified in an interactive terminal, xlflow shows a Bubble Tea selector. JSON and non-interactive runs must pass `--agent` or `--target`.

`xlflow run` accepts repeatable typed arguments with the `string:`, `int:`, and `bool:` prefixes. `string:` may carry an empty value. Successful runs report `macro.duration_ms` in JSON and print the elapsed time plus save behavior in plain text. Failing runs return `macro_failed` with VBA error metadata including module name, `Err.Number`, `Err.Description`, and `error.line` when VBA exposes `Erl`, and the non-JSON message stays readable as `Main Err 5: inputPath is required` or `Main line 10 Err 5: inputPath is required` when a line number is available. The default run does not save the workbook; use `--save` or `--save-as` explicitly.

`xlflow trace inject [workbook]` injects the opt-in `XlflowTrace` VBA module. After injection, VBA code can call `XlflowLog "message"` or `Call XlflowLog("message")`; `xlflow run --trace` initializes a temp log file, runs the macro, and returns trace events in the top-level JSON `trace` field.

`xlflow test` discovers argument-free VBA `Sub` procedures from the configured workbook when their names start with `Test` or end with `_Test`. New and initialized projects include `src/modules/XlflowAssert.bas`, which provides scalar-only `AssertEquals expected, actual, [message]`. Compare object properties such as `Range.Value2`, not object references.

`xlflow diff` compares two `.xlsx`/`.xlsm`/`.xltx`/`.xltm` workbooks with excelize and reports sheet, cell value, and formula differences. With `--vba-before` and `--vba-after`, it also compares exported `.bas`, `.cls`, and `.frm` source trees as normalized text. Differences are reported as successful results; use `diff.summary.total_diffs` in JSON to decide whether anything changed.

```vb
Public Sub TestCreateReport()
    AssertEquals 10, Sheets("Result").Range("A1").Value
End Sub
```

## Verification

Run `task verify` as the obvious local automated verification entry point. It is fast and deterministic, and currently executes non-COM test coverage via `go test ./...`.

```bash
task verify
```

For Excel COM E2E verification, run on a Windows environment with Excel and VBIDE access enabled, then confirm `doctor --json` reports healthy Excel/VBIDE diagnostics.

Use the repo-local `xlflow-tmp-workspace-e2e` skill and run the `tmp_workspaces` flow as the standard real-workbook path. Success criteria are: `new/doctor/pull/lint` all succeed, then module/class/form round-trip checks succeed via `push/run/test/pull/lint`.
