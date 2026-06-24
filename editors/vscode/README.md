# xlflow for VS Code

This extension adds VS Code support for source-controlled Excel VBA projects powered by `xlflow`.

The extension is a thin client for `xlflow lsp --stdio` and the xlflow CLI. Diagnostics, hover, completion, signature help, symbols, definition and reference lookup, formatting, and VBA/COM type inference are provided by the Go-based xlflow language server.

## Requirements

- Install `xlflow`.
- Make `xlflow` available on `PATH`, or set `xlflow.path` to the executable path.

## Development

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
- `xlflow.completion.triggerSuggestInStatements`: trigger VS Code suggestions in likely VBA statement contexts.
- `xlflow.completion.progIdsInStrings`: trigger VS Code suggestions inside `CreateObject("...")` and `GetObject("...")` strings.

## Commands

The command palette includes:

- `xlflow: Restart Language Server`
- `xlflow: Check Environment`
- `xlflow: New Project`
- `xlflow: Initialize Project`
- `xlflow: Pull Workbook`
- `xlflow: Push Sources`
- `xlflow: Run Macro`
- `xlflow: Run Tests`
- `xlflow: Lint Workspace`
- `xlflow: Format Document`
- `xlflow: Format Project`
- `xlflow: Save Workbook`
- `xlflow: Start Session`
- `xlflow: Session Status`
- `xlflow: Stop Session`

Workbook commands run from the resolved workspace folder. `New Project` runs `xlflow new`, `Initialize Project` runs `xlflow init <workbook>`, `Pull Workbook` runs `xlflow pull`, `Push Sources` runs `xlflow push`, `Run Macro` runs `xlflow run`, `Run Tests` runs `xlflow test`, `Lint Workspace` runs `xlflow lint`, `Format Project` runs `xlflow fmt --write`, and `Save Workbook` runs `xlflow save`.

`Format Document` invokes VS Code document formatting for the active editor. For VBA files, formatting is provided by `xlflow lsp --stdio`.

Session commands run `xlflow session start`, `xlflow session status`, and `xlflow session stop` from the resolved workspace folder.

## Testing

The extension registers a VS Code Test Explorer controller for VBA tests. Discovery runs `xlflow test list --json` and execution runs `xlflow test --json --module <module> --filter <name>` from the selected workspace folder. The TypeScript extension does not parse VBA or generate test cases itself.

`xlflow: Run Tests` remains available as a command palette escape hatch that runs `xlflow test` and writes the CLI output to the `xlflow` output channel.

## Output

Use the `xlflow` output channel for CLI command output and language client messages. Use `xlflow Language Server Trace` for LSP trace output.

## Known Limitations

- The extension does not install or bundle `xlflow`.
- Macro selection is not interactive yet; `xlflow: Run Macro` runs the configured default macro.
- There are no webviews, workbook previews, or rich Excel session management UI.
- `xlflow: New Project` and `xlflow: Initialize Project` expose only the base CLI workflow, without option pickers for `--with-skill`, `--with-module`, `--agent`, or `--json`.
- The extension does not implement VBA parsing, diagnostics, formatting, completion candidates, symbol analysis, or type inference in TypeScript.
