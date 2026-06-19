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

The MVP server supports full document synchronization, diagnostics, document symbols, workspace symbols, definition lookup, and hover. Open editor buffers are authoritative over saved filesystem content until the editor sends `textDocument/didClose`.

The built-in VBA/COM database includes practical Excel, MSForms, Scripting, ADODB, VBIDE, Office, and VBA constant metadata for hover and basic type inference.

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
