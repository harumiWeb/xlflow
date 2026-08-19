# ADR-0022: Conservative VBA Control-Flow Graph

## Status

Accepted

## Context

ADR-0021 establishes a protocol-neutral, procedure-level VBA analysis IR. Its
statement nesting is syntactic: it cannot answer whether a statement is
reachable, whether an assignment holds on every path, or whether cleanup runs
before every exit. Issue #427 needs those answers as a shared foundation for
the reliability analysis proposed by issue #425.

VBA control flow includes structured branches and loops, unstructured labels
and `GoTo`, procedure and loop exits, process-terminating `End`, and path-local
`On Error` modes. Malformed source, duplicate or missing labels, and dynamic
`Resume` targets also prevent a graph builder from always identifying one exact
successor. Treating those cases as impossible would let guarantee-oriented
rules report unsafe conclusions.

The graph must serve both batch analysis and immutable LSP snapshots without
depending on CLI or LSP protocol types. ADR-0023 adds interprocedural effect
summaries as a separate consumer layer; the CFG remains conservative and does
not depend on that information.

## Decision

Add `internal/vba/cfg` as a separate layer built from
`procedureir.DocumentIR`. Do not embed graph state in `ProcedureIR` or retain
tree-sitter values. Each procedure graph uses deterministic, source-ordered
block and edge IDs and supports defensive cloning.

Represent normal versus exceptional flow independently from certainty. An edge
is classified by its control-flow meaning and as normal or exceptional, and it
may separately be marked uncertain. This allows, for example, both a known
normal successor and an uncertain exceptional successor from one statement
without conflating their semantics.

Each graph has synthetic entry, normal-exit, exceptional-exit,
termination-exit, and unknown-exit blocks. Structured branches, selection,
loops, nearest-loop exits, procedure exits, labels, `GoTo`, `On Error`,
`Resume`, and standalone `End` map to explicit transitions. Missing, duplicate,
or recovered targets must not be resolved to an arbitrary statement.

Model the active VBA error mode with forward, path-sensitive dataflow.
`On Error GoTo <label>`, `On Error Resume Next`, and `On Error GoTo 0` update
that mode along their paths and are merged conservatively. Every executable
statement remains a possible fault site even after ADR-0023 adds effect
summaries. The builder therefore adds uncertain exceptional transitions to the
active handler, the resume-next successor, or the exceptional exit as
appropriate. Dynamic or recovered `Resume` behavior and other unresolved
control flow pass through conservative unknown flow rather than disappearing.

Guarantee-oriented queries include uncertain edges. Reachability,
unreachability, dominance, definite assignment, exit transitions, and
all-paths-cleanup queries must therefore avoid proving a property by omitting an
unresolved path. Consumers may request an explicit view, such as normal flow
without exceptional edges, when the rule's contract calls for it; uncertainty
within that selected view still participates.

Each immutable graph revision owns reusable query indexes for statement-to-block
lookup and incoming/outgoing edge traversal. The index also retains the
conservative reachability sets for the supported edge-filter views. Graph value
copies may share these read-only indexes, while defensive clones and graph
transformations rebuild them together with their copied or changed slices.
Graphs constructed directly by internal callers without an index use a
correctness-preserving fallback index. Indexed graph revisions treat the block
and edge slices, `Entry`, `UnknownExit`, and `UnknownFlowSources` as immutable;
replacing either slice or changing any reachability input on a copied graph
causes the index validity guard to rebuild the index, while content changes
must go through defensive cloning or a graph transformation.

Cache the document CFG on the immutable `AnalysisSnapshot`, using the same
ownership, concurrency, retryable-cancellation, and defensive-copy contract as
procedure IR. Editor snapshots build it lazily from cached IR; batch snapshots
may seed it with an already-built CFG for the exact immutable revision. A
successor in the same lifecycle may reuse a
completed procedure graph fragment when both its procedure source and module
semantic context match. The module-context component is required because CFG
assignment facts contain resolved symbol scopes. Source ranges are rebased to
the current procedure while graph and block IDs remain deterministic within
that procedure. Recovery ambiguity, conditional-compilation changes, module
context changes, and close/reopen force safe rebuilds.

Migrate `VBA204` as the first CFG consumer. It detects handler fallthrough by
normal-flow reachability of the handler label rather than by scanning preceding
text. Its diagnostic ID, fields, ordering, cleanup-label exception, inline
suppression, CLI JSON, and LSP ranges remain unchanged. No public CLI option,
configuration key, JSON shape, diagnostic ID, or LSP capability is added.

The same graph construction also owns protocol-neutral validation facts for
procedure-local control-flow legality. The procedure IR carries source ranges,
normalized loop variables, `Next` variable lists, and conditional/recovery
markers; the CFG resolves label definitions and active structured-block stacks
once and records facts for duplicate/undefined labels, `Next` mismatches, and
invalid exits. Facts retain statement IDs, ranges, expected/actual values, and
a certainty bit, but never diagnostic IDs or presentation text. Facts are
copied and rebased with graph cache artifacts. Any parser recovery,
conditional-compilation ambiguity, or unresolved nesting makes the fact
uncertain and the diagnostic projector omits it (fail-open to syntax or
unknown-flow diagnostics).

`internal/vba/cfg.ValidationDiagnostics` is the single projector for these
facts. Lint and compile-equivalent analyze adapters use it directly, while
LSP consumes the lint projection or the same projector for batch findings and
converts byte ranges to UTF-16 only at the protocol boundary. This keeps
`VB055`--`VB058` parity and prevents a second control-flow parser from being
introduced in an adapter.

## Consequences

- Positive: batch and LSP reliability checks share one deterministic
  control-flow model instead of implementing rule-specific walkers.
- Positive: normal, exceptional, termination, and unresolved outcomes remain
  distinguishable for later analyses.
- Positive: conservative uncertainty prevents guarantee queries from producing
  false confidence on malformed or dynamically targeted control flow.
- Positive: the graph remains protocol-neutral and safe after parser retirement
  because it contains only Go-owned IR values and ranges.
- Positive: repeated analyzer queries reuse immutable lookup, adjacency, and
  reachability data instead of rebuilding maps or scanning all edges.
- Negative: treating every executable statement as a potential fault site adds
  edges and can reduce precision; ADR-0023 does not narrow this fallback.
- Negative: path-sensitive error modes and conservative unknown flow make graph
  construction and query testing more complex than a structured-only CFG.
- Negative: each built graph retains additional index memory, and transformed
  graphs must rebuild the index when their edge set changes.
- Negative: callers must choose a flow view deliberately; using a normal-only
  view for a general guarantee would be unsound.
- Limitation: graph IDs are stable only for the same document revision and are
  not persistent workspace identities.

## Alternatives Considered

1. **Build CFG data into `ProcedureIR`** - Rejected because syntax ownership and
   executable-path analysis evolve independently, and consumers that need only
   normalized syntax should not pay for or depend on graph construction.
2. **Construct a separate CFG inside each rule** - Rejected because label
   resolution, error modes, recovery, and exit semantics would drift between
   batch and LSP consumers.
3. **Ignore exceptional flow until effect analysis exists** - Rejected because
   guarantee queries would incorrectly treat faulting paths as impossible.
4. **Treat unresolved targets as unreachable** - Rejected because missing,
   duplicate, recovered, and dynamic targets are uncertainty, not proof that no
   execution path exists.
5. **Mix exceptional flow and uncertainty into one edge kind** - Rejected
   because a transition can be exceptional but precisely known, or normal but
   uncertain; consumers need both dimensions.
6. **Reuse graph fragments by source hash alone** - Rejected because module
   declarations can change resolved scopes without changing a procedure body.
   Reuse therefore also requires the module semantic context and the existing
   generation-safe snapshot lifecycle.

## Evidence

- Requirements: xlflow issues #425 and #427.
- Procedure IR boundary and recovery contract:
  `docs/adr/ADR-0021-procedure-analysis-ir.md`,
  `docs/specs/vba-analysis-ir.md`.
- Analyzer ownership: `docs/adr/ADR-0013-analyze-runtime-risk-ownership.md`,
  `internal/analyze/analyzer.go`.
- Snapshot and LSP boundary:
  `docs/adr/ADR-0014-reusable-vba-lsp-server.md`,
  `internal/vba/intel/analysis_snapshot.go`.
- Existing validation behavior: `internal/analyze/analyzer_test.go`.

## Related

- `docs/adr/ADR-0013-analyze-runtime-risk-ownership.md`
- `docs/adr/ADR-0014-reusable-vba-lsp-server.md`
- `docs/adr/ADR-0021-procedure-analysis-ir.md`
- `docs/adr/ADR-0023-procedure-effect-summaries.md`
- `docs/specs/vba-control-flow-graph.md`
- xlflow issues #425, #426, #427, and #428
