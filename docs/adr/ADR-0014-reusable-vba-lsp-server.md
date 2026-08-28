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
language server entry point, with document-kind dispatch for additional
xlflow-owned source formats.

- Keep CLI flag handling in `internal/cli`.
- Keep LSP protocol handling, JSON-RPC stdio transport, URI conversion, and
  protocol type conversion in `internal/lspserver`.
- Keep VBA source analysis in protocol-neutral packages under `internal/vba`.
- Classify documents before analysis. VBA documents retain tree-sitter snapshots;
  UserForm specifications under the configured `src.forms/specs` root retain raw
  source and use protocol-neutral YAML syntax helpers under
  `internal/excel/forms/intel`. Unrelated YAML and JSON documents are ignored.
- Keep the practical VBA/COM metadata database in `internal/vbadb`.
- Represent analysis results with xlflow-owned structures such as `Range`,
  `Diagnostic`, `Symbol`, `Location`, and `Hover`; convert them to LSP protocol
  structs only in `internal/lspserver`.
- Treat open LSP documents as authoritative over filesystem content until
  `didClose`.
- Advertise incremental document synchronization and apply ranged changes in
  client order using LSP UTF-16 positions. Retain full-document replacements as
  a compatibility fallback, and replace only the changed document's immutable
  analysis snapshot. Refresh semantic tokens when the open workspace changes,
  because their classification can depend on project symbols from other files.
- Reuse a previous tree-sitter tree only by cloning it under the previous
  snapshot's serialized tree lease, editing that clone with byte-based
  tree-sitter coordinates, and parsing the new source with a parser local to
  that operation. The published previous tree is never edited.
- Let each immutable snapshot own its parsed tree. Retiring a superseded
  snapshot rejects new readers and closes its tree only after active readers
  complete. Candidate snapshots publish only when their captured document
  generation and lifecycle still match the open document.
- Keep `didOpen` and `didChange` non-blocking with respect to derived analysis.
  They synchronously register the current immutable snapshot, reserve
  its document generation, and schedule background workspace-overlay and
  diagnostics work; they do not wait for that work to finish.
- Split workspace preparation into an interactive declaration index and a
  background semantic index. The declaration layer owns project symbols and
  exact/prefix/qualified/module/kind lookup and becomes ready independently;
  ProcedureIR, CFG, call sites, and project-resolution artifacts remain behind
  the semantic readiness boundary. Initial declaration work uses a fixed,
  bounded worker pool, while semantic workers leave capacity for interactive
  requests and are cancellable during shutdown.
- Publish open-document declarations and semantic artifacts independently but
  under the same generation. A pending or incomplete semantic entry cannot
  make project-wide negative conclusions complete, while current open-document
  declarations remain available from the immutable document snapshot.
- While an open document's workspace overlay is pending, mask both its saved
  filesystem entry and its previously published overlay. Workspace-wide queries
  may therefore omit that document temporarily, but must not return stale
  symbols or calls. Document-local requests continue to read the current
  immutable snapshot directly.
- Permit only the latest matching document generation, lifecycle, and snapshot
  to publish derived workspace or diagnostic results. Superseding changes,
  close, reopen, and shutdown cancel obsolete work; canceled derived-artifact
  builds remain retryable and are never cached as completed results.
- After a change, publish procedure-scoped Fast diagnostics after the 300 ms
  edit debounce, then build the matching workspace overlay and replace the Fast
  result with Full diagnostics after two seconds of editor idle. A Full result
  outranks Fast for the same generation, and neither phase may publish across a
  generation or lifecycle boundary. Opening a small or non-VBA document keeps
  its existing Full-diagnostics path. Opening a VBA document with at least
  10,000 source lines publishes a bounded procedure-local Fast preview
  immediately. After that publication the cold overlay starts, while the Full
  pass is released by the existing open-delay timer.
- Anchor reusable procedure diagnostics to a procedure identity and source
  fragment. Fast publication may rebase unchanged procedure-local diagnostics
  to the current procedure start, but must omit interprocedural results until
  Full analysis revalidates them and must reapply current inline suppressions.
- Build one immutable workspace-resolution view per diagnostic request. Pending
  overlays remain masked in the workspace index while current open-document
  declarations are merged into that request view without synchronous overlay
  publication.
- On unreconcilable edit coordinates or invalid version ordering, retain the
  last valid snapshot until a later full-document replacement can resynchronize
  it. When a valid new source cannot use an old tree, parse it completely and
  record the fallback in opt-in performance logging.
- Load a curated built-in database for practical Excel, MSForms, Scripting,
  ADODB, VBIDE, Office, and VBA constant/type metadata.

### Amendment: snapshot-scoped interactive symbol indexes (Issue #756)

Each immutable `AnalysisSnapshot` owns one source-oriented interactive index
for its revision. The index serves exact, qualified, and prefix symbol lookup,
current-procedure and module declaration lookup, and the symbol-kind filtering
needed by interactive completion and signature help. Interactive consumers
reuse the snapshot's procedure catalog and index instead of rebuilding a
whole-document symbol model for each request.

Normal hover, definition, completion, and signature-help requests must be able
to use these source declarations without building ProcedureIR, CFG, or Full
diagnostics. A successor revision receives a new immutable index; no postings
from an older revision are observable. Compatibility callers that do not have
an analysis snapshot may retain their conservative fallback.

The opt-in performance recorder exposes stable structural counters for index
builds and hits, procedure-catalog builds and reuse, exact/prefix/qualified
queries, full document-symbol builds, and interactive full-symbol fallbacks.
These counters are developer telemetry only and do not alter LSP responses,
symbol visibility, ordering, or cancellation and generation safety.

### Amendment: end-to-end startup critical path (Issue #757)

The extension must not make language intelligence wait for a separate
`xlflow version` preflight or for project/sidebar detail refreshes. The
language client starts optimistically after the workspace and LSP configuration
metadata needed to construct its document selector are available. Process
spawn or initialization failure remains actionable through the existing
availability and restart paths; the short availability timeout is not applied
to the complete LSP initialization because TypeLib preparation can be a valid
cold-start cost.

Startup timing is opt-in developer telemetry. A client-generated anonymous ID
correlates client lifecycle records with server lifecycle and readiness records
through the child process environment. Monotonic elapsed times are interpreted
within each process, with wall-clock timestamps used only for correlation. The
telemetry does not add protocol fields or requests, and a non-empty hover,
definition, or completion response is required before reporting a first usable
language feature. The CLI invokes best-effort TypeLib preparation through the
server's pre-construction hook after the server baseline is captured, so that
cold preparation remains part of the measured startup interval.

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
- Positive: Opening or changing a very large VBA module does not make document
  lifecycle notifications wait for full workspace and diagnostic analysis.
- Positive: Cross-file declaration queries can become useful before the
  workspace has built every file's IR, CFG, and call-site artifacts.
- Positive: Snapshot-scoped source declarations let latency-sensitive
  interactive requests reuse one immutable index across hover, definition,
  completion, and signature help.
- Negative: The server coordinates two readiness and publication lifecycles;
  semantic consumers must continue to fail open until both declaration and
  semantic completeness are proven.
- Negative: Initial indexing uses bounded parallel source work and therefore
  needs explicit cancellation and immutable-result ownership at shutdown.
- Negative: The main binary now carries LSP protocol and JSON-RPC dependencies.
- Negative: URI, path normalization, and UTF-16 position conversion become part
  of xlflow's long-lived compatibility surface.
- Negative: The document store must retain source and a line-offset index, and
  reject malformed or out-of-order edits without corrupting editor state.
- Negative: Incremental parsing needs temporary cloned trees and eagerly parses
  changed revisions; this trades a small per-edit allocation for safely
  reusing unchanged syntax structure without exposing mutable trees to readers.
- Negative: Workspace-wide queries are conservatively incomplete while an open
  document's newest overlay is still being built, and results become complete
  only after background publication succeeds.
- Negative: Each supported non-VBA document kind needs an explicit analyzer
  adapter and must not fall through to VBA symbols, semantic tokens, or edits.
- Negative: The curated COM database requires maintenance until a TypeLib
  importer and patch pipeline are available.
- Negative: Every immutable document revision retains an additional index, and
  compatibility callers without a snapshot still need an explicit fallback
  path.

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
- `docs/adr/ADR-0033-vba-error-outcomes-and-lsp-project-cache.md`
