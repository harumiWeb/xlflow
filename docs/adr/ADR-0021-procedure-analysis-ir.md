# ADR-0021: Shared Procedure-Level VBA Analysis IR

## Status

Accepted

## Context

xlflow's VBA consumers already share the owned tree-sitter parser, but they do
not share one normalized representation of a procedure. `internal/analyze`,
`internal/vba/calls`, `internal/vba/symbols`, and `internal/vba/intel` each
derive overlapping procedure, declaration, call, range, and event facts for
their own use. Adding control-flow and interprocedural reliability analysis
under issue #425 would multiply those traversals and make CLI and LSP behavior
more likely to drift.

Raw tree-sitter nodes are unsuitable as the shared contract. They are valid
only during `ast.ParsedDocument.Read`, their tree is serialized and owned by the
parsed document, and an LSP snapshot can retire that document while older
analysis values remain in use. LSP protocol types are also unsuitable because
the same analysis must remain reusable by source-only CLI commands.

The foundation in issue #426 needs exact existing diagnostic locations,
recoverable partial analysis, deterministic output, and project resolution
that can be refreshed without rewalking unchanged syntax. It does not yet need
control-flow graphs or effect propagation. Issue #699 adds a performance
constraint: batch resolution diagnostics must not deep-clone all procedure
syntax and semantic payloads when only project-dependent facts change.

## Decision

Add a protocol-neutral, procedure-level intermediate representation under
`internal/vba/procedureir`.

The IR builder performs one traversal inside `ast.ParsedDocument.Read` and
copies only Go-owned values: strings, enums, stable IDs, slices, and
`ast.Range`. It never retains tree-sitter nodes, trees, or source slices from
the read callback. `BuildSource` owns parsing and closing for one-shot callers;
`BuildParsed` uses a caller-owned parsed document without closing it.

Separate syntax extraction from project resolution. A syntax-local, immutable
`DocumentIR` records procedures, declarations, statements, expressions, calls,
and variable accesses. `ResolveView` applies a replaceable project
symbol/call resolver to a read-only resolution overlay without reparsing,
rescanning source, or cloning the procedure payloads. The view exposes resolved
calls, accesses, and events through helpers so consumers do not depend on
whether facts are stored inline or in a side table. Resolution must explicitly
represent matched, ambiguous, unresolved, external, built-in-like, and
member-call outcomes.

The overlay is indexed by revision-local procedure identity and stable IR fact
IDs, not source text. A procedure identity includes document/module context,
qualified procedure name and kind, declaration start, and source-order
ordinal. Call IDs are reused for calls; accesses and event references receive
the corresponding revision-local IDs. These IDs are deterministic within one
document revision but are not workspace-global identities. Access and event
IDs remain internal metadata and are omitted from serialized IR/JSON.

`DiagnosticsView` shares the existing resolution-diagnostic walker with
`Diagnostics` and reads syntax/recovery facts from the immutable IR while
reading project-dependent call, access, and event facts from the overlay. Batch
analyzer resolution diagnostics use this view. Effects, ordinary analyzer
rules, lint, metrics, and developer-only oracle consumers may retain the
materialized path when they require an independently owned document. The LSP
project snapshot instead keeps the canonical syntax-local `DocumentIR` and
attaches a `ResolvedDocumentView`; its capability, effect, realtime,
dependency, and resolution-diagnostic consumers read the view, with only
bounded per-procedure fact projections where a legacy loop requires
value-shaped calls or accesses. The normal LSP path does not call
`Resolve`/`Materialize`.

Keep `Resolve` as the compatibility API for consumers that require an
independently owned resolved `DocumentIR`: it may materialize a view and must
preserve the existing input-unchanged and independent-ownership guarantees.
`ResolveView` and materialization must also preserve behavior for a nil resolver,
including the existing unresolved/incomplete representation.

When one project revision needs both procedure-only and compile-equivalent
resolution, construct one canonical symbol index and derive the two resolver
semantics as immutable views. The procedure-only view filters the canonical
procedure entries at lookup time instead of owning a second full candidate map.
Document resolution views may share revision-local fact IDs; each mode retains
its own overlay so the intentional semantic difference is never merged. A
canceled or incomplete project-resolution build is not published to the
revision cache and remains retryable.

Carry module kind in resolver symbols. A receiver-less call may resolve to a
non-standard procedure only within the same module; class, document, and
UserForm procedures require an explicit receiver across module boundaries.

Use `ast.Range` as the canonical source coordinate contract. LSP UTF-16
conversion remains in the LSP adapter; the IR does not import LSP protocol
types.

Make the immutable `AnalysisSnapshot` the ownership and cache boundary for
editor and batch consumers. An editor snapshot lazily constructs syntax IR
from its owned parsed document, while a batch snapshot may seed the same cache
with an already-built IR for the exact immutable revision. Both paths support
concurrent readers and return defensive copies. A successor in the same
open-document lifecycle may inherit completed Go-owned procedure fragments
keyed by procedure kind/name/ordinal, exact source hash, and module-context
hash. Ranges are rebased from the fragment declaration start and document-wide
declaration IDs are deterministically reassigned when materialized. In-flight,
canceled, panicked, recovered, or ambiguous fragments are not inherited;
close/reopen always starts with an empty artifact store.

Centralize procedure-kind and event-handler classification in the IR. `Sub`,
`Function`, and `Property Get/Let/Set` bodies are procedures. A module-level
`Event Foo(...)` declaration is a declaration, not a procedure body.
Document-module `Workbook_*` and `Worksheet_*` procedures, `Auto_Open` and
`Auto_Close`, and non-test UserForm procedures with a non-empty
`<source>_<event>` shape carry event metadata compatible with the existing
`internal/vba/intel` behavior.

Malformed but parseable source produces partial IR with document recovery flags
and recovered or unknown statements where useful facts remain available.
Consumers retain authority over whether recovery is acceptable; in particular,
batch `analyze` may continue to reject parser recovery even though the builder
does not discard the partial IR.

Issue #426 models syntactic statement nesting only. Basic blocks, control-flow
edges, reachability, and path queries belong to the CFG layer in ADR-0022.
Direct/transitive effect summaries and fixed-point propagation belong to the
separate effects layer in ADR-0023.

Migrate `VBA204` error-handler fallthrough detection as the validation consumer.
Its diagnostic fields, ordering, cleanup-label exception, inline suppression,
CLI JSON, and LSP ranges remain unchanged. Existing `inspect calls` output is
also a compatibility contract; projecting call facts from the IR must not
change its JSON shape or resolution meaning.

Analyzer procedure projections are revision-local immutable views over the
canonical `DocumentIR`/`ProcedureIR` storage. They may retain pointers to
owned IR and CFG elements, plus compact facts and analysis overlays, but must
not copy complete procedure collections or retain parser leases, tree-sitter
nodes, source buffers, or obsolete revisions. Internal analyzer readers use
read-only views/iterators; snapshot and public compatibility boundaries keep
their existing defensive-copy behavior. This decision is limited to analyzer
projection storage and does not change the separate `Resolve` or CFG ownership
contracts.

## Consequences

- Positive: lint, analyze, inspect, LSP, and later project-wide analyses can
  reuse one procedure and source-location model.
- Positive: syntax extraction is independent of changing project symbols, so
  call resolution can be refreshed without another CST walk.
- Positive: read-only batch resolution can refresh project-dependent facts
  without deep-cloning statements, expressions, declarations, calls, accesses,
  and other procedure-level payloads.
- Positive: project resolution modes share the canonical symbol storage and
  revision-local fact identity, reducing duplicate index and preparation
  allocations while preserving their separate semantics.
- Positive: cached analysis values are safe after the parsed document closes
  because the IR contains no borrowed parser state.
- Positive: a body-only edit rebuilds only changed procedure IR while safely
  rebasing unchanged fragments after preceding line insertions or deletions.
- Positive: partial syntax facts remain available for editor scenarios while
  strict batch consumers can continue to fail closed.
- Positive: CFG and effect work get an explicit, narrow input boundary instead
  of expanding rule-specific CST walkers.
- Negative: the normalized model must evolve when tree-sitter-vba adds syntax
  that cannot be represented by current statement or expression kinds.
- Negative: revision-local analyzer views require every consumer and worker to
  honor the immutable ownership contract; mutable rule state must remain local.
- Negative: defensive copies and normalization allocate more Go values than a
  consumer-specific walk.
- Negative: compatibility projections remain necessary while existing calls,
  symbols, analyzer, and LSP result types continue to serve their public
  contracts.
- Negative: overlay indexes and revision-local fact identity add bookkeeping,
  and consumers that need independent ownership still pay the materialization
  cost of `Resolve`.
- Negative: procedure-only lookup performs a filtered view projection over the
  canonical candidate slice; callers must keep resolver entries immutable so
  the view cannot leak mutations back into the full resolver.
- Limitation: issue #426 provides syntactic structure, not executable-path
  truth, type-complete member binding, COM type-library resolution, or
  interprocedural effects.

## Alternatives Considered

1. **Continue traversing raw CST nodes in every rule** - Rejected because
   procedure discovery, ranges, calls, and recovery handling would continue to
   diverge, and later CFG/effect work would duplicate parser-lifetime logic.
2. **Retain tree-sitter nodes in a shared model** - Rejected because nodes are
   borrowed from a serialized, closeable `ParsedDocument` and cannot safely
   outlive its `Read` callback or owning snapshot.
3. **Create a separate model per analyzer rule** - Rejected because it preserves
   duplicated extraction and gives later analyses no common deterministic
   input.
4. **Make the shared model LSP-specific** - Rejected because CLI analysis and
   inspection must not depend on protocol coordinates or lifecycle types, and
   ADR-0014 already confines LSP conversion to the adapter boundary.
5. **Build CFG and effects into the first IR** - Rejected because those layers
   require separate graph and propagation decisions covered by issues #427 and
   #428. Keeping the foundation syntactic makes #426 independently testable.
6. **Rebuild syntax whenever the project symbol snapshot changes** - Rejected
   because local syntax is unchanged; a separate resolution overlay is cheaper
   and makes unresolved or ambiguous states explicit.
7. **Deep-clone the complete IR for every read-only diagnostic pass** - Rejected
   because large modules duplicate syntax-local and procedure-level semantic
   collections even though batch resolution changes only a small set of facts.
8. **Key overlay entries by source text** - Rejected because text is not a
   stable identity when the IR already provides deterministic revision-local
   fact IDs and procedure identity.
9. **Migrate every materialized consumer at once** - Rejected because LSP,
   effects, and other consumers may require independently owned snapshots; the
   overlay is introduced first at the batch resolution-diagnostic boundary.

## Evidence

- Requirements: xlflow issues #425 and #426.
- Existing parser lifetime and range contract:
  `internal/vba/ast/ast.go`, `internal/vba/ast/ast_test.go`.
- Existing call extraction and resolution contract:
  `internal/vba/calls/calls.go`, `internal/vba/calls/calls_test.go`.
- Existing snapshot ownership and incremental parsing:
  `internal/vba/intel/analysis_snapshot.go`,
  `internal/vba/intel/analysis_snapshot_test.go`.
- Validation consumer: `internal/analyze/analyzer.go`,
  `internal/analyze/analyzer_test.go`.
- Resolution overlay and compatibility boundary: `internal/vba/procedureir`,
  `internal/analyze/analyzer.go`, and issue #699.
- Public compatibility contract: `docs/specs/cli-contract.md`.

## Related

- `docs/adr/ADR-0013-analyze-runtime-risk-ownership.md`
- `docs/adr/ADR-0014-reusable-vba-lsp-server.md`
- `docs/adr/ADR-0022-conservative-vba-control-flow-graph.md`
- `docs/adr/ADR-0023-procedure-effect-summaries.md`
- `docs/adr/ADR-0048-revision-scoped-semantic-query-dag.md`
- `docs/specs/vba-analysis-ir.md`
- xlflow issues #425, #426, #427, #428, and #699
