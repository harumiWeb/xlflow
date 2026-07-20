# xlflow lsp

Start the reusable VBA language server for editor integrations.

## Usage

```bash
xlflow lsp --stdio
xlflow lsp --stdio --log-file .xlflow/lsp.log
xlflow lsp --stdio --performance-log
xlflow lsp --check
xlflow lsp --version
```

## Options and Arguments

| Option / argument   | Description                                                 | Default |
| ------------------- | ----------------------------------------------------------- | ------- |
| `--stdio`           | Run the LSP server over standard input/output JSON-RPC.     | false   |
| `--check`           | Validate server dependencies without starting JSON-RPC.     | false   |
| `--version`         | Print LSP server build metadata.                            | false   |
| `--log-file`        | Write server logs to this file instead of standard error.   | stderr  |
| `--performance-log` | Log structured performance measurements for LSP operations. | false   |

Exactly one of `--stdio`, `--check`, or `--version` is required.

## Notes

`xlflow lsp --stdio` reserves stdout exclusively for LSP messages. Normal logs go to stderr unless `--log-file` is provided.

Performance logging is opt-in. With `--performance-log`, each measured operation writes one structured log line to stderr or the configured `--log-file`. Records include the operation, document URI or path where applicable, version or generation, input bytes and lines, elapsed time, result count, and outcome. Diagnostics records also identify discarded obsolete results, and cache-aware operations may report a cache hit or miss.

`xlflow lsp --check` works even before a project has an `xlflow.toml`; it validates the parser and built-in type database using default configuration.

When the global generated TypeLib database is missing or stale, `xlflow lsp --stdio` and `xlflow lsp --check` attempt best-effort generation before loading the database. Generation failures are reported on stderr and do not prevent the LSP from starting with the embedded built-in database.

The server advertises incremental document synchronization and applies ordered UTF-16 ranged edits, including CRLF and supplementary Unicode source. For ranged edits it clones and edits the previous tree-sitter tree with UTF-8 byte coordinates before parsing the new source; the previously published tree remains immutable for in-flight requests. It retains whole-document replacements as a compatibility fallback. Invalid coordinates or version ordering preserve the last valid buffer until a full replacement resynchronizes it. With performance logging enabled, each change records whether parsing was incremental, a full fallback, or retained and includes the fallback reason when applicable. It also supports diagnostics, semantic tokens, document symbols, workspace symbols, definition lookup, references, prepare rename, rename, hover, completion, signature help, document formatting, and CodeLens. Open editor buffers are authoritative over saved filesystem content until the editor sends `textDocument/didClose`.

Rename support is conservative and scope-aware. It works for high-confidence VBA source symbols such as local variables, parameters, procedure-local constants, private module-level variables/constants, same-module private procedures, and same-procedure labels used by `GoTo`, `GoSub`, `Resume`, or `On Error GoTo`. It refuses host object model members, TypeLib/external members, public project-wide APIs, module files, `Attribute VB_Name`, UserForm controls, event handlers, property groups, ambiguous names, and unresolved identifiers.

Diagnostics reuse xlflow's file-local VBA lint rules against the current in-memory editor buffer and publish stable `VB...` codes with `source="xlflow"`. Opening a document starts diagnostics immediately. Changes are debounced and coalesced so a document never has more than one active diagnostics worker, and only the newest open version can publish results. Slow documents do not prevent other documents from being analyzed. Closing a document cancels its pending work and publishes a final empty diagnostics result; an older analysis cannot publish afterward. Project-wide and filesystem-only lint checks remain available through `xlflow lint`.

Direct YAML files under the configured `src.forms/specs` directory are also validated against the shared UserForm contract while editing. Syntax failures use `UFY001`; contract errors and support-level warnings use stable `UFV001`–`UFV014` codes with `source="xlflow"`. The server points unknown or unsupported properties at their keys, invalid values and references at their values, and missing or structural problems at the nearest affected YAML node. Files in that directory are treated as UserForm candidates even with an invalid `kind`, so `kind` itself can be diagnosed; unrelated YAML files remain ignored. JSON specifications are detected but do not yet provide editor diagnostics.

UserForm YAML completion uses that same contract. It offers root, form, and control fields; built-in control types and matching ProgIDs; fixed values and booleans; and eligible same-document `parentId` references. Control properties are filtered by the selected type, and completion remains useful while a normal YAML edit is temporarily incomplete. Snapshot-oriented fields appear only after their name is typed. Document and control-entry snippets are also available; JSON specifications do not yet provide completion.

For an enabled `VB044` procedure-name constant mismatch, diagnostics are published for the current editor buffer on open and after edits. `textDocument/codeAction` offers a `quickfix` that replaces only the direct string literal with the enclosing procedure name. This does not add missing constants or change procedure rename behavior.

The LSP also publishes editor-first argument diagnostics such as `VB030` for missing required arguments, excessive arguments, and unknown named arguments when the target signature is known from project symbols or the built-in VBA/COM database. These diagnostics are intentionally LSP-only for now and are not yet part of the `xlflow lint` CLI contract.

The built-in VBA/COM database includes practical Excel, MSForms, Scripting, ADODB, VBIDE, Office, and VBA constant metadata for hover, completion, and basic type inference.

Semantic tokens are provided by the Go language server with full-document `textDocument/semanticTokens/full` responses and `textDocument/semanticTokens/full/delta` updates. They classify VBA declarations, parameters, variables, built-in types, globals, constants, member expressions, comments, strings, numbers, operators, and keywords. Full responses carry opaque result IDs; the server retains the four most recent results for each open document and returns a delta only when its JSON payload is smaller than a full response. Unknown, expired, cross-document, or closed-document IDs fall back to a full response. Range semantic token requests are not advertised.

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

Set `xlflow.lsp.performanceLogging` to `true` to have the VS Code extension pass `--performance-log`. The setting defaults to `false`.

## Related

- [lint](./lint)
- [inspect symbols](./inspect)
- [JSON Output](../reference/json-output)
