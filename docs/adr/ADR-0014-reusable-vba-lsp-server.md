# ADR-0014: Reusable VBA LSP Server Boundary

## Status

Accepted

## Context

xlflow needs editor-grade VBA intelligence for diagnostics, symbols, hover,
completion, references, and definition lookup. The same analysis must also
remain available to CLI commands such as `lint`, `analyze`, `inspect symbols`,
and future agent integrations.

Placing parsing logic inside a VS Code extension would create a second VBA
analysis stack and make CLI and editor behavior drift. Placing LSP protocol
types directly inside the VBA packages would make those packages harder to reuse
from non-LSP commands.

The LSP server also has a strict transport contract: stdout must contain only
framed JSON-RPC messages while the server is running. Normal command logging and
preflight diagnostics therefore need explicit separation from the stdio
transport.

## Decision

Implement `xlflow lsp --stdio` in the main xlflow binary as the reusable VBA
language server entry point.

- Keep CLI flag handling in `internal/cli`.
- Keep LSP protocol handling, JSON-RPC stdio transport, URI conversion, and
  protocol type conversion in `internal/lspserver`.
- Keep VBA source analysis in protocol-neutral packages under `internal/vba`.
- Keep the practical VBA/COM metadata database in `internal/vbadb`.
- Represent analysis results with xlflow-owned structures such as `Range`,
  `Diagnostic`, `Symbol`, `Location`, and `Hover`; convert them to LSP protocol
  structs only in `internal/lspserver`.
- Treat open LSP documents as authoritative over filesystem content until
  `didClose`.
- Advertise incremental document synchronization and apply ranged changes in
  client order using LSP UTF-16 positions. Retain full-document replacements as
  a compatibility fallback, and replace only the changed document's immutable
  analysis snapshot.
- Load a curated built-in database for practical Excel, MSForms, Scripting,
  ADODB, VBIDE, Office, and VBA constant/type metadata.

The VS Code extension should remain a thin language client that launches:

```ts
{
  command: "xlflow",
  args: ["lsp", "--stdio"]
}
```

## Consequences

- Positive: CLI, editor, and future agent features can share the same VBA
  analysis behavior.
- Positive: LSP dependencies stay isolated from the reusable VBA analysis layer.
- Positive: The server can be launched by any editor or tool that supports LSP
  stdio transport.
- Positive: Unsaved editor buffers can be diagnosed and queried without writing
  temporary source files.
- Negative: The main binary now carries LSP protocol and JSON-RPC dependencies.
- Negative: URI, path normalization, and UTF-16 position conversion become part
  of xlflow's long-lived compatibility surface.
- Negative: The document store must retain source and a line-offset index, and
  reject malformed or out-of-order edits without corrupting editor state.
- Negative: The curated COM database requires maintenance until a TypeLib
  importer and patch pipeline are available.

## Alternatives Considered

1. **Implement LSP inside `xlflow-vscode`** - Rejected because it duplicates
   parser, linter, symbol, and resolver behavior outside the core project.
2. **Expose LSP protocol types from `internal/vba`** - Rejected because CLI and
   agent callers should not depend on LSP structs.
3. **Hand-roll a dependency-free LSP implementation** - Rejected for the MVP
   because existing protocol and JSON-RPC packages are sufficient when confined
   to the adapter layer.
4. **Start with a TypeLib importer instead of a curated database** - Rejected for
   the MVP because redistribution, patching, and completeness policy need a
   separate design. A curated database provides immediate hover and inference
   utility.

## Related

- `docs/specs/cli-contract.md`
- `internal/lspserver`
- `internal/vba/intel`
- `internal/vbadb`
- `vitepress/commands/lsp.md`
