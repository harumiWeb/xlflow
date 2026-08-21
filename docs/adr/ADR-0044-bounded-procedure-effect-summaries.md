# ADR-0044: Bounded Procedure-Effect Fixed Points and Indexed CFG Queries

## Status

Accepted

## Context

ADR-0023 and ADR-0033 require deterministic, provenance-bearing procedure and
error summaries. Their compatibility projections keep direct and propagated
evidence, uncertainty, error evidence, and representative call chains
separate. That contract is correct, but an implementation that eagerly copies
all of those collections through every confirmed call edge makes the fixed
point state grow with the number of transitive origins. Long chains, wide
fan-in, and dense project-local graphs can therefore spend most of their time
allocating duplicate evidence and call-chain slices even after the semantic
answer has stopped changing.

Issue #674 (parent #670) also identified a remaining graph-query risk: a
reachable-block query must not scan every graph block for every reachable block
when the CFG already has stable block and statement indexes.

The optimization must not alter diagnostic meaning. Direct versus propagated
effects, error handling outcomes, uncertainty, source evidence, diagnostic
call paths, effect origins, IDs, locations, messages, multiplicity, and stable
ordering are compatibility requirements. A real-world corpus delta is an
observation to review, not permission to weaken a rule.

## Decision

Keep the existing `ProjectSummary` and `ProcedureSummary` compatibility
projection. Internally, build the fixed point from two layers:

1. a compact semantic state whose membership is bounded by the finite effect,
   error-outcome, and uncertainty domains; and
2. minimal direct evidence and deterministic witness seeds used to explain a
   fact when a consumer asks for provenance.

Confirmed caller/callee adjacency is deduplicated and indexed. A caller is
re-evaluated only when a semantic fact, uncertainty member, or required
witness seed becomes new. A procedure's direct facts remain direct; facts
arriving across a confirmed edge remain propagated. Recursive and diamond
graphs therefore converge without repeatedly growing equivalent transitive
origin state.

Detailed diagnostic evidence and representative call paths are materialized
from the retained witness seeds and the stable call graph when needed. Error
witnesses use one replay of the legacy global procedure worklist per shared
`ProjectSummary` materialization cache; an isolated per-origin search is not
allowed because another origin can requeue a branch and affect the first path
that reaches an owner. The materialized result uses the same
procedure/edge/source ordering as the pre-optimization projection. If a
consumer requires an origin, uncertainty, or error path, the implementation
must retain enough seed information to reconstruct it deterministically; it
may not replace a required witness with an arbitrary representative.

The effects package exposes `BuildWithStats(documents)`, returning the same
`ProjectSummary` as `Build` plus developer-facing `BuildStats`. `Build` remains
the compatibility wrapper. The stats report `WorklistEvaluations`,
`MaxPropagatedFactsPerProcedure`, and `TotalPropagatedFacts`; these counters
refer to compact semantic facts, not the size of a lazily materialized
provenance projection.

CFG consumers use the shared block-ID and outgoing-edge indexes. Reachable
block-to-statement lookup is an indexed average-O(1) operation and does not
perform a block-by-block scan. Graph copies and transformations preserve or
rebuild index validity according to the existing CFG contract.

The analyzer may expose the effect counters through the existing opt-in
`analyze --performance-log` stage record. The counters are observability only:
they do not enter analyzer JSON, LSP diagnostics, finding ordering, exit
codes, or configuration semantics.

## Consequences

- Large call graphs keep a bounded semantic fixed-point state and avoid
  repeated transitive evidence allocation.
- Diagnostic consumers retain the existing evidence and path contract, with
  provenance work paid when the consumer actually needs it.
- Worklist and propagated-fact counters make scaling regressions measurable at
  500, 1,000, and 2,000-procedure workloads and on real corpus projects.
- The lazy witness boundary adds reconstruction code and requires regression
  coverage for direct/propagated distinctions, uncertainty, recursion,
  diamonds, and representative error paths.
- The semantic domain remains intentionally conservative. Unresolved,
  ambiguous, external, and dynamic calls do not become confirmed effects just
  because provenance is compacted.

## Alternatives Considered

1. **Keep eagerly copying every transitive collection** - Rejected because
   fixed-point memory and allocation costs grow with graph connectivity rather
   than with the finite semantic domain.
2. **Propagate booleans only and discard witnesses** - Rejected because
   `VBA203`, `VBA220`, `VBA237`, and `VBA244` require source evidence,
   uncertainty, or deterministic paths.
3. **Reconstruct paths by an unconstrained graph search at every diagnostic** -
   Rejected because it could change representative-path selection and make
   output depend on traversal order. Reconstruction must use stable indexes and
   the existing ordering contract.
4. **Remove propagated evidence from the compatibility projection** -
   Rejected because internal consumers and regression tests rely on the
   direct/propagated distinction even though the fixed-point implementation is
   compact.
5. **Retain the reachable-block linear scan** - Rejected because the graph
   already owns the indexes needed to answer the query without quadratic work.

## Evidence

- Requirements: xlflow issues #670 and #674.
- Existing effect contract and consumers:
  `docs/adr/ADR-0023-procedure-effect-summaries.md`,
  `docs/specs/vba-effect-analysis.md`, and
  `internal/vba/effects`.
- Existing error-outcome and path contract:
  `docs/adr/ADR-0033-vba-error-outcomes-and-lsp-project-cache.md` and
  `docs/specs/vba-error-propagation-analysis.md`.
- CFG indexes and reachability semantics:
  `docs/specs/vba-control-flow-graph.md` and `internal/vba/cfg`.
- Performance instrumentation and corpus procedure:
  `docs/specs/cli-contract.md` and
  `docs/specs/static-analysis-corpus.md`.

## Supersedes

None. This ADR amends the implementation strategy recorded by ADR-0023 and
ADR-0033 without changing their diagnostic contracts.

## Superseded by

None.

## Related

- `docs/adr/ADR-0023-procedure-effect-summaries.md`
- `docs/adr/ADR-0033-vba-error-outcomes-and-lsp-project-cache.md`
- `docs/adr/ADR-0035-procedure-call-cycle-analysis.md`
- `docs/adr/ADR-0043-bounded-batch-analysis-parallelism.md`
- xlflow issues #670 and #674
