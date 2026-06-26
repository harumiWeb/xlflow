# xlflow for VS Code

This extension adds VS Code support for source-controlled Excel VBA projects powered by `xlflow`.

The extension is a thin client for `xlflow lsp --stdio` and the xlflow CLI. Diagnostics, hover, completion, signature help, symbols, definition and reference lookup, formatting, CodeLens, and VBA/COM type inference are provided by the Go-based xlflow language server.

## Requirements

- Install `xlflow`.
- Make `xlflow` available on `PATH`, or set `xlflow.path` to the executable path.

## Sidebar

The extension contributes an `xlflow` Activity Bar container with native TreeViews. It does not use a Webview.

When the selected workspace folder does not contain `xlflow.toml`, the sidebar shows setup actions only:

- `New Project`
- `Init Existing Workbook`
- `Run Doctor`
- `Open Documentation`

When `xlflow.toml` exists, the sidebar switches to project mode:

- `Project`: workspace, configured workbook, `xlflow.toml`, session state, and save-required state.
- `Modules`: standard, class, document, and UserForm modules discovered from `xlflow inspect symbols --json`.
- `Tests`: tests discovered from `xlflow test list --json`, with shortcuts for run all and single-test execution.

Project view title actions refresh state, pull workbook source, push source changes, and toggle the managed session. `Push Sources` asks for confirmation before running.

## Development

Use Node.js 22 or newer. The extension test runner uses `@vscode/test-electron` 3.x.

From this directory:

```bash
pnpm install
pnpm compile
```

To launch the extension in VS Code development mode, open this folder and run the extension host from the Run and Debug view after compiling.

## Settings

Configure the executable path when `xlflow` is not on `PATH`:

```json
{
  "xlflow.path": "C:\\path\\to\\xlflow.exe"
}
```

Common settings:

- `xlflow.lsp.enabled`: start `xlflow lsp --stdio` for VBA files.
- `xlflow.lsp.logFile`: log file passed to the language server. The default is `.xlflow/lsp.log`.
- `xlflow.lsp.trace.server`: trace verbosity for the language server trace output channel.
- `xlflow.codeLens.enabled`: show xlflow CodeLens actions above runnable VBA procedures.
- `xlflow.codeLens.runProcedure`: show `Run` actions above runnable VBA procedures.
- `xlflow.codeLens.runTests`: show `Run Test` actions above VBA test procedures.
- `xlflow.codeLens.userFormEvents`: show `Run` actions above UserForm event handlers.
- `xlflow.run.saveBeforeRun`: save dirty VBA documents before running a procedure from CodeLens.
- `xlflow.completion.triggerSuggestInStatements`: trigger VS Code suggestions in likely VBA statement contexts.
- `xlflow.completion.progIdsInStrings`: trigger VS Code suggestions inside `CreateObject("...")` and `GetObject("...")` strings.
- `xlflow.testing.autoDiscover`: automatically discover VBA tests when an xlflow workspace opens.

## Commands

The command palette includes:

- `xlflow: Restart Language Server`
- `xlflow: Check Environment`
- `xlflow: New Project`
- `xlflow: Initialize Project`
- `xlflow: Install Agent Skill`
- `xlflow: Install Helper Modules`
- `xlflow: Pull Workbook`
- `xlflow: Push Sources`
- `xlflow: Run Macro`
- `xlflow: Run Procedure`
- `xlflow: Run Test Procedure`
- `xlflow: Run Tests`
- `xlflow: Lint Workspace`
- `xlflow: Format Document`
- `xlflow: Format Project`
- `xlflow: Save Workbook`
- `xlflow: Start Session`
- `xlflow: Session Status`
- `xlflow: Restart Session`
- `xlflow: Stop Session`
- `xlflow: Open Output`
- `xlflow: Refresh Project`
- `xlflow: Refresh Modules`
- `xlflow: Refresh Tests`
- `xlflow: Run All Tests`
- `xlflow: Run Doctor`
- `xlflow: Toggle Session`
- `xlflow: Open Documentation`

Workbook commands run from the resolved workspace folder. `New Project` runs `xlflow new`, `Initialize Project` runs `xlflow init <workbook>`, `Install Agent Skill` runs `xlflow skill install --agent <provider>`, `Install Helper Modules` runs `xlflow module install` or `xlflow module install --push`, `Pull Workbook` runs `xlflow pull`, `Push Sources` runs `xlflow push`, `Run Macro` runs `xlflow run`, `Run Tests` runs `xlflow test`, `Lint Workspace` runs `xlflow lint`, `Format Project` runs `xlflow fmt --write`, and `Save Workbook` runs `xlflow save`.

`Install Agent Skill` prompts for one of the bundled provider targets: `codex`, `claude`, `cursor`, `gemini`, or `agents`. It also asks whether to pass `--force` before replacing an existing skill installation. `Install Helper Modules` prompts before using `--push` because that mode imports the helper modules into the configured workbook.

The language server supplies CodeLens actions for no-argument VBA `Sub` procedures. `$(play) Run` invokes `xlflow run <qualifiedName>`, and `$(beaker) Run Test` invokes `xlflow --json test --module <moduleName> --filter <name>`. VS Code renders the `$(...)` prefixes as codicons. Dirty VBA documents are saved first when `xlflow.run.saveBeforeRun` is enabled.

`Format Document` invokes VS Code document formatting for the active editor. For VBA files, formatting is provided by `xlflow lsp --stdio`.

Session commands run `xlflow session start`, `xlflow --json session status`, `xlflow session stop`, and restart from the resolved workspace folder.

## Session Status

The extension shows a lightweight xlflow session indicator in the Status Bar:

- `$(circle-slash) xlflow: No Project`: the selected workspace folder has no `xlflow.toml`.
- `$(circle-slash) xlflow: No Session`: no active session.
- `$(check) xlflow: Session Active`: an active session is available.
- `$(sync~spin) xlflow: Starting` or `$(sync~spin) xlflow: Stopping`: session start or stop is running.
- `$(warning) xlflow: Session Error`: session status or operation failed.

Click the Status Bar item to start, stop, restart, inspect the session, open the output channel, or run `xlflow doctor`. In setup mode, it opens setup actions instead. Active sessions use a green status color. Session details and command output are written to the `xlflow` output channel.

## Testing

The extension registers a VS Code Test Explorer controller for VBA tests. Discovery runs `xlflow test list --json` and execution runs `xlflow test --json --module <module> --filter <name>` from the selected workspace folder. The TypeScript extension does not parse VBA or generate test cases itself.

When `xlflow.testing.autoDiscover` is enabled, startup discovery runs only for workspace folders that contain `xlflow.toml`. Manual Test Explorer refresh remains available regardless of this setting.

`xlflow: Run Tests` remains available as a command palette escape hatch that runs `xlflow test` and writes the CLI output to the `xlflow` output channel.

## Output

Use the `xlflow` output channel for CLI command output and language client messages. Use `xlflow Language Server Trace` for LSP trace output.

## Known Limitations

- The extension does not install or bundle `xlflow`.
- Macro selection is not interactive yet; `xlflow: Run Macro` runs the configured default macro. Runnable no-argument `Sub` procedures can be launched from CodeLens.
- The sidebar is native TreeView UI only. There are no webviews, workbook previews, or rich HTML dashboards.
- `xlflow: New Project` and `xlflow: Initialize Project` expose only the base CLI workflow, without option pickers for `--with-skill`, `--with-module`, `--agent`, or `--json`.
- The extension does not implement VBA parsing, diagnostics, formatting, completion candidates, symbol analysis, or type inference in TypeScript.
