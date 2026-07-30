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
control-flow graphs or effect propagation.

## Decision

Add a protocol-neutral, procedure-level intermediate representation under
`internal/vba/procedureir`.

The IR builder performs one traversal inside `ast.ParsedDocument.Read` and
copies only Go-owned values: strings, enums, stable IDs, slices, and
`ast.Range`. It never retains tree-sitter nodes, trees, or source slices from
the read callback. `BuildSource` owns parsing and closing for one-shot callers;
`BuildParsed` uses a caller-owned parsed document without closing it.

Separate syntax extraction from project resolution. A syntax-local
`DocumentIR` records procedures, declarations, statements, expressions, calls,
and variable accesses. `Resolve` applies a replaceable project symbol/call
resolver to a copy of that IR without reparsing or rescanning source. Resolution
must explicitly represent matched, ambiguous, unresolved, external,
built-in-like, and member-call outcomes.

Use `ast.Range` as the canonical source coordinate contract. LSP UTF-16
conversion remains in the LSP adapter; the IR does not import LSP protocol
types.

Make the immutable `AnalysisSnapshot` the LSP ownership and cache boundary.
Each snapshot lazily constructs at most one syntax IR from its owned parsed
document, supports concurrent readers, and returns defensive copies. A new
incrementally parsed snapshot gets its own IR; issue #426 does not splice
procedure IR fragments across revisions.

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
edges, reachability, and path queries belong to issue #427. Direct/transitive
effect summaries and fixed-point propagation belong to issue #428.

Migrate `VBA204` error-handler fallthrough detection as the validation consumer.
Its diagnostic fields, ordering, cleanup-label exception, inline suppression,
CLI JSON, and LSP ranges remain unchanged. Existing `inspect calls` output is
also a compatibility contract; projecting call facts from the IR must not
change its JSON shape or resolution meaning.

## Consequences

- Positive: lint, analyze, inspect, LSP, and later project-wide analyses can
  reuse one procedure and source-location model.
- Positive: syntax extraction is independent of changing project symbols, so
  call resolution can be refreshed without another CST walk.
- Positive: cached analysis values are safe after the parsed document closes
  because the IR contains no borrowed parser state.
- Positive: partial syntax facts remain available for editor scenarios while
  strict batch consumers can continue to fail closed.
- Positive: CFG and effect work get an explicit, narrow input boundary instead
  of expanding rule-specific CST walkers.
- Negative: the normalized model must evolve when tree-sitter-vba adds syntax
  that cannot be represented by current statement or expression kinds.
- Negative: defensive copies and normalization allocate more Go values than a
  consumer-specific walk.
- Negative: compatibility projections remain necessary while existing calls,
  symbols, analyzer, and LSP result types continue to serve their public
  contracts.
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
- Public compatibility contract: `docs/specs/cli-contract.md`.

## Related

- `docs/adr/ADR-0013-analyze-runtime-risk-ownership.md`
- `docs/adr/ADR-0014-reusable-vba-lsp-server.md`
- `docs/specs/vba-analysis-ir.md`
- xlflow issues #425, #426, #427, and #428
