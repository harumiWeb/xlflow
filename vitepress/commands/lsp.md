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

The MVP server supports full document synchronization, diagnostics, document symbols, workspace symbols, definition lookup, references, hover, and completion. Open editor buffers are authoritative over saved filesystem content until the editor sends `textDocument/didClose`.

Diagnostics reuse xlflow's file-local VBA lint rules against the current in-memory editor buffer and publish stable `VB...` codes with `source="xlflow"`. Project-wide and filesystem-only lint checks remain available through `xlflow lint`.

The built-in VBA/COM database includes practical Excel, MSForms, Scripting, ADODB, VBIDE, Office, and VBA constant metadata for hover, completion, and basic type inference.

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
