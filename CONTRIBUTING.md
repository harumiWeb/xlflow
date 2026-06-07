# Contributing to xlflow

Thank you for your interest in contributing to xlflow.

xlflow is a weekend/side project, so responses to issues, discussions, and pull requests may be slow. That said, contributions are very welcome. Bug reports, feature ideas, documentation improvements, tests, real-world feedback, and pull requests are all appreciated.

## Project goal

xlflow aims to make Excel VBA development more CLI-first and agent-friendly.

The core idea is to provide a safe proof loop for Excel VBA work:

```text
pull → edit → push → lint → test/run → save
```

When contributing, please keep the following goals in mind:

- Make Excel VBA easier to develop from the CLI
- Make workflows safe and understandable for AI agents
- Prefer deterministic JSON output for automation
- Avoid UI-driven or modal workflows when possible
- Preserve user workbooks and source files safely
- Keep failure messages actionable

## Ways to contribute

You can help in many ways:

- Report bugs
- Suggest features
- Improve documentation
- Improve examples
- Add or improve tests
- Test xlflow with real Excel workbooks
- Improve Windows / Excel / VBIDE compatibility
- Improve AI agent workflow guidance
- Improve error messages and JSON output

Real-world feedback is especially valuable because Excel VBA projects vary widely across environments, language settings, workbook formats, and corporate security policies.

## Before opening an issue

Before reporting a bug, please check the following when possible:

```bash
xlflow doctor --json
xlflow lint --json
xlflow macros --json
```

For runtime failures, structured debug output is often useful:

```bash
xlflow run <MacroName> --json
```

When opening an issue, please include:

- OS version
- Excel version
- xlflow version or commit SHA
- The command you ran
- Full JSON output if available
- Whether VBIDE access is enabled
- A minimal reproduction if possible

Please avoid sharing confidential workbooks or company data. If a workbook contains sensitive information, try to reduce it to a minimal sample before attaching anything.

## Development setup

Clone the repository and run xlflow from source:

```bash
git clone https://github.com/harumiWeb/xlflow.git
cd xlflow
go run ./cmd/xlflow --help
```

Install locally:

```bash
go install ./cmd/xlflow
```

For development and CI, treat the Go version declared in `go.mod` as the supported toolchain source of truth. Repository CI and release workflows resolve Go from that file rather than duplicating a version string elsewhere.

`go install` may contact the Go module mirror and checksum database configured in your Go environment. Interactive `xlflow new` and `xlflow init` may also query the latest GitHub Release to render an update notice in the scaffold welcome banner; use `--no-update-check` or `XLFLOW_NO_UPDATE_CHECK=1` when you need to suppress that network call.

Repository linting uses `golangci-lint` and `PSScriptAnalyzer`. Ensure `Invoke-ScriptAnalyzer` is available in your PowerShell environment before running the lint task or pre-commit hook.

```bash
task lint
```

Run the fast verification task:

```bash
task verify
```

The fast verification path currently runs non-COM test coverage via:

```bash
go test ./...
```

Run vulnerability and third-party licence inventory checks with:

```bash
task verify:security
```

Excel COM E2E verification should be done on Windows with Microsoft Excel installed and **Trust access to the VBA project object model** enabled.

## Release preflight

Before a release that touches workbook automation, VBA import/export, run/test behavior, session handling, or other Excel COM paths, do not rely on CI alone.

Run the repo-local `xlflow-tmp-workspace-e2e` skill against fresh `tmp_workspaces` and verify at least:

- blank workbook scaffold: `new`, `doctor`, `pull`, `lint`
- standard module round-trip: `push`, `run`, workbook-state verification, `pull`, `lint`
- class module round-trip
- UserForm round-trip including `.frm` and `.frx`
- `init` from an existing workbook

If the release changes session behavior, also verify the session loop:

- `session start`
- `push --fast --session --no-save`
- `run --session` and/or `test --session`
- `save --session`
- `session stop`

## Pull request guidelines

Pull requests are welcome.

For small fixes, feel free to open a PR directly. For larger changes, opening an issue first is helpful so we can discuss the direction before implementation.

A good PR should usually include:

- A clear summary of the change
- The motivation for the change
- Tests or verification notes
- Documentation updates when behavior changes
- Screenshots or command output when useful

Please keep PRs focused. Smaller PRs are easier to review and merge.

## Coding guidelines

### Go code

- Keep command behavior explicit and predictable
- Prefer clear error messages over clever abstractions
- Preserve stable JSON contracts when possible
- Avoid changing exit-code semantics without updating documentation
- Keep Excel COM-dependent behavior isolated from testable non-COM logic when practical
- Add tests for argument validation, output contracts, and non-COM behavior

Run formatting and tests before submitting:

```bash
go fmt ./...
go test ./...
```

### CLI and JSON behavior

AI agents and scripts rely on xlflow's JSON output. When changing command behavior, consider whether the change affects:

- `status`
- `command`
- `error.code`
- `error.message`
- command-specific fields such as `macro`, `macros`, `tests`, `diff`, `debug`, or `issues`
- exit codes

Human-readable output can evolve more freely, but machine-readable JSON should remain stable whenever possible.

### VBA-related behavior

xlflow should encourage VBA that works well in unattended automation.

Prefer workflows that avoid:

- `Select`
- `Activate`
- `ActiveSheet` assumptions
- File picker dialogs
- `InputBox`
- Modal `MsgBox`
- Broad `On Error Resume Next`

When adding or changing generated VBA support modules, keep them small, predictable, and easy to inspect.

## Documentation

Documentation contributions are welcome.

Useful documentation improvements include:

- Clearer setup steps for Windows and Excel
- Troubleshooting for VBIDE access
- Examples of `pull` / `push` / `run` / `test` workflows
- AI agent usage examples
- Real-world VBA project migration notes
- Better explanations of JSON output and exit codes

If you find something confusing, that is a good documentation issue or PR.

## Testing with real Excel workbooks

Real workbook testing is very valuable.

If you test xlflow on a workbook and find a problem, please describe:

- Workbook type: `.xlsm`, `.xlsx`, `.xltm`, etc.
- Whether it contains standard modules, class modules, UserForms, or document modules
- Whether it uses non-ASCII module names or Japanese text
- Whether it contains `.frx` UserForm companion files
- Which xlflow command failed
- Whether `doctor --json` succeeds

Please do not upload confidential workbooks. A minimized sample workbook is best.

## Security

Please do not report security-sensitive issues in a public issue if they could expose users or workbooks to risk.

If a dedicated security contact is added in the future, please use that route. Until then, avoid posting sensitive exploit details publicly and open a minimal issue asking for a private contact path.

## Maintainer availability

xlflow is maintained as a weekend/side project. Review and response times may vary, and some issues may take time to investigate, especially when they require a specific Windows, Excel, COM, or corporate environment.

Even if responses are slow, contributions and reports are welcome and appreciated.

## License

By contributing to xlflow, you agree that your contributions will be licensed under the MIT License.
