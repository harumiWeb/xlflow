# xlflow lsp

Start the reusable VBA language server for editor integrations.

## Usage

```bash
xlflow lsp --stdio
xlflow lsp --stdio --log-file .xlflow/lsp.log
xlflow lsp --check
xlflow lsp --version
```

## Options and Arguments

| Option / argument | Description                                               | Default |
| ----------------- | --------------------------------------------------------- | ------- |
| `--stdio`         | Run the LSP server over standard input/output JSON-RPC.   | false   |
| `--check`         | Validate server dependencies without starting JSON-RPC.   | false   |
| `--version`       | Print LSP server build metadata.                          | false   |
| `--log-file`      | Write server logs to this file instead of standard error. | stderr  |

Exactly one of `--stdio`, `--check`, or `--version` is required.

## Notes

`xlflow lsp --stdio` reserves stdout exclusively for LSP messages. Normal logs go to stderr unless `--log-file` is provided.

`xlflow lsp --check` works even before a project has an `xlflow.toml`; it validates the parser and built-in type database using default configuration.

When the global generated TypeLib database is missing or stale, `xlflow lsp --stdio` and `xlflow lsp --check` attempt best-effort generation before loading the database. Generation failures are reported on stderr and do not prevent the LSP from starting with the embedded built-in database.

The MVP server supports full document synchronization, diagnostics, semantic tokens, document symbols, workspace symbols, definition lookup, references, prepare rename, rename, hover, completion, signature help, document formatting, and CodeLens. Open editor buffers are authoritative over saved filesystem content until the editor sends `textDocument/didClose`.

Rename support is conservative and scope-aware. It works for high-confidence VBA source symbols such as local variables, parameters, procedure-local constants, private module-level variables/constants, same-module private procedures, and same-procedure labels used by `GoTo`, `GoSub`, `Resume`, or `On Error GoTo`. It refuses host object model members, TypeLib/external members, public project-wide APIs, module files, `Attribute VB_Name`, UserForm controls, event handlers, property groups, ambiguous names, and unresolved identifiers.

Diagnostics reuse xlflow's file-local VBA lint rules against the current in-memory editor buffer and publish stable `VB...` codes with `source="xlflow"`. Project-wide and filesystem-only lint checks remain available through `xlflow lint`.

The LSP also publishes editor-first argument diagnostics such as `VB030` for missing required arguments, excessive arguments, and unknown named arguments when the target signature is known from project symbols or the built-in VBA/COM database. These diagnostics are intentionally LSP-only for now and are not yet part of the `xlflow lint` CLI contract.

The built-in VBA/COM database includes practical Excel, MSForms, Scripting, ADODB, VBIDE, Office, and VBA constant metadata for hover, completion, and basic type inference.

Semantic tokens are provided by the Go language server with full-document `textDocument/semanticTokens/full` responses. They classify VBA declarations, parameters, variables, built-in types, globals, constants, member expressions, comments, strings, numbers, operators, and keywords. Range and delta semantic token requests are not advertised yet.

CodeLens is provided through `textDocument/codeLens` with `resolveProvider=false`. The server returns `$(play) Run` for runnable no-argument `Sub` procedures and `$(beaker) Run Test` for no-argument test procedures named `Test*`, `test*`, or `*_Test`; VS Code renders those `$(...)` prefixes as codicons. Function, Property, Declare, and argument-bearing procedures are intentionally excluded. UserForm event-style procedures are hidden by default unless the editor passes `initializationOptions.codeLens.userFormEvents=true`.

For UserForms, xlflow reads tracked `.frm` files and extracts design-time controls for code intelligence. Controls such as `txtName` in `Me.txtName.Text` and `Me.Controls("txtName").Text` can participate in hover, completion, and definition lookup without opening Excel.

## VS Code Client

The VS Code extension should launch xlflow as a thin language client:

```ts
serverOptions = {
  command: "xlflow",
  args: ["lsp", "--stdio"],
};
```

For the MVP, xlflow is resolved from `PATH`.

## Related

- [lint](./lint)
- [inspect symbols](./inspect)
- [JSON Output](../reference/json-output)
