# xlflow LSP Dev Client

This is a minimal VS Code Language Client used only for local development and manual verification of `xlflow lsp --stdio`.

This is not the production `xlflow-vscode` extension.

The production extension should be implemented separately once the LSP protocol surface stabilizes.

## Setup

Install dependencies from this directory:

```powershell
pnpm install
```

Compile the extension:

```powershell
pnpm compile
```

For type-check only:

```powershell
pnpm check
```

## xlflow on PATH

The development client launches `xlflow` from `PATH` and does not bundle an xlflow binary.

From the repository root, install the current local xlflow build with:

```powershell
task install
```

Confirm that VS Code can resolve it:

```powershell
xlflow lsp --check
```

## Running in VS Code

From the repository root:

1. Open the repository root in VS Code.
2. Select `Run xlflow LSP Dev Client` in Run and Debug.
3. Press `F5` or choose `Run and Debug: Start Debugging`.
4. In the Extension Development Host, open an xlflow project or a folder containing VBA source.
5. Open a `.bas`, `.cls`, or `.frm` file.

The root launch configuration uses:

```text
--extensionDevelopmentPath=${workspaceFolder}/tools/vscode-lsp-dev
```

This prevents VS Code from treating the repository root package as the extension manifest.

The prelaunch task runs `pnpm --dir tools/vscode-lsp-dev compile`; on Windows it uses `pnpm.cmd` directly to avoid shell argument parsing issues.

If you open `tools/vscode-lsp-dev` directly in VS Code, run `pnpm compile` before pressing `F5`.

The client starts:

```text
xlflow lsp --stdio --log-file .xlflow/lsp.log
```

The server working directory is the first VS Code workspace folder in the Extension Development Host.

## Syntax highlighting

This dev client contributes a VBA TextMate grammar for `.bas`, `.cls`, and `.frm` files.

The grammar is intentionally useful enough to serve as a reference for a future production extension. It highlights:

- comments, strings, date literals, numeric literals, and line continuations
- `Attribute` lines and VBA preprocessor directives
- procedures, properties, declares, enums, types, variables, parameters, and labels
- VBA control-flow keywords, storage modifiers, primitive types, common Excel/COM types, and common constants
- member access such as `.Range` / `.Value`

The grammar is still development-only and should be treated as a smoke/reference implementation, not production branding or packaging.

## Logs

Use VS Code's Output panel:

- `xlflow LSP Dev Client` for client lifecycle logs.
- `xlflow LSP Dev Client Trace` for Language Client trace logs. Set the channel log level to Trace when debugging protocol traffic.

The server log file is written under the opened workspace:

```text
.xlflow/lsp.log
```

Standard output is reserved for LSP JSON-RPC frames. Non-LSP server logs should appear in stderr or `.xlflow/lsp.log`, not stdout.

## Manual verification

- [ ] Open this extension in VS Code
- [ ] Start Extension Development Host
- [ ] Open an xlflow project or VBA source folder
- [ ] Open a `.bas`, `.cls`, or `.frm` file
- [ ] Confirm VBA syntax highlighting is applied to comments, strings, declarations, keywords, and member access
- [ ] Confirm `(` and `"` insert their closing pair in VBA files
- [ ] Confirm `xlflow lsp --stdio` starts
- [ ] Confirm no non-LSP logs are printed to stdout
- [ ] Confirm server logs are written to `.xlflow/lsp.log`
- [ ] Confirm diagnostics appear in Problems when server-side diagnostics are available
- [ ] Confirm diagnostics clear after fixing the source
- [ ] Confirm signature help appears after `(`, `,`, and a space in parenless calls
- [ ] Confirm argument diagnostics appear for missing required arguments such as `dict.Add "A"` and `rng.Find()`
- [ ] Confirm document symbols appear in Outline when implemented
- [ ] Confirm hover works when implemented
- [ ] Confirm go to definition works when implemented
- [ ] Confirm Japanese paths do not break startup
- [ ] Confirm Japanese comments / string literals do not break ranges

Useful smoke source:

```vb
Option Explicit

Sub Smoke()
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "A",
    dict.Add "A"

    Dim rng As Range
    rng.Find(What:="A", LookAt:=
    rng.Find()
    rng.Font.Co

    MsgBox "Hello",
End Sub
```

Expected checks:

- `dict.Add "A",` shows `Item` as the active parameter.
- `dict.Add "A"` reports a missing required argument.
- `rng.Find(What:="A", LookAt:=` shows `LookAt` as the active parameter.
- `rng.Find()` reports a missing `What` argument.
- `rng.Font.Co` offers `Color`.
- `MsgBox "Hello",` shows `Buttons` as the active parameter.
