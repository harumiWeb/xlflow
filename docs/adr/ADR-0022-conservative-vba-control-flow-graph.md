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

Cache the document CFG lazily on the immutable `AnalysisSnapshot`, using the
same revision ownership, concurrency, error caching, and defensive-copy
contract as the procedure IR cache. An incremental edit creates a new cache and
does not reuse graph fragments or IDs across revisions.

Migrate `VBA204` as the first CFG consumer. It detects handler fallthrough by
normal-flow reachability of the handler label rather than by scanning preceding
text. Its diagnostic ID, fields, ordering, cleanup-label exception, inline
suppression, CLI JSON, and LSP ranges remain unchanged. No public CLI option,
configuration key, JSON shape, diagnostic ID, or LSP capability is added.

## Consequences

- Positive: batch and LSP reliability checks share one deterministic
  control-flow model instead of implementing rule-specific walkers.
- Positive: normal, exceptional, termination, and unresolved outcomes remain
  distinguishable for later analyses.
- Positive: conservative uncertainty prevents guarantee queries from producing
  false confidence on malformed or dynamically targeted control flow.
- Positive: the graph remains protocol-neutral and safe after parser retirement
  because it contains only Go-owned IR values and ranges.
- Negative: treating every executable statement as a potential fault site adds
  edges and can reduce precision; ADR-0023 does not narrow this fallback.
- Negative: path-sensitive error modes and conservative unknown flow make graph
  construction and query testing more complex than a structured-only CFG.
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
6. **Reuse graph fragments across incremental revisions** - Rejected because
   statement and graph IDs are revision-local, and reuse would complicate
   snapshot lifetime and stale-publication guarantees.

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
