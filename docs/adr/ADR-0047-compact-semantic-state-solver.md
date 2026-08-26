# ADR-0047: Compact Semantic State Solver for Procedure Data Flow

## Status

Accepted

## Context

The procedure analysis wave has already moved expensive setup behind immutable
procedure facts, applicability plans, semantic kernel result stores, and
independent HTTP/generic data-flow lanes. The remaining HTTP and Array kernels
still carry most of their propagation state as string-keyed maps. Every block
transfer clones those maps, Array joins allocate a union-key map and copy shape
slices, and HTTP state equality depends on a deep comparison of nested maps.

On the ROneCOne `analyze-only` workload at the Issue #713 starting revision
(`6c9f8ba60c5b162cb7115e8e68744412c7de9d5d`), a Windows amd64 profile run on
the repository's i7-12700 development host used
`rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./internal/staticanalysis/corpus -run '^$' -bench '^BenchmarkRealWorldCorpus/ronecone/analyze-only$' -benchmem -benchtime=1x -count=1 -timeout=10m` with the test binary and heap profile retained locally under `%TEMP%\xlflow-issue713-baseline.*`. The focused
alloc-space view attributed approximately 1.62 GB to `cloneArrayState`,
0.53 GB to `meetArrayState`, and 0.07 GB to `cloneHTTPState` in that sample.
These copies are semantic state, not diagnostic evidence, and their cost grows
with the width of a procedure and the number of CFG joins.

This decision relates to the ownership and identity boundary in ADR-0021, the
CFG identity and conservative edge contract in ADR-0022, the bounded semantic
state/evidence separation in
`docs/adr/ADR-0044-bounded-procedure-effect-summaries.md`, and the
kernel/result/work-budget contract in ADR-0046. It does not move the generic
source-to-sink package,
change CFG storage, or introduce a persistent cache.

## Decision

Add an internal `internal/analyze/semanticstate` solver boundary used first by
the HTTP and Array kernels.

### Ownership and identity

Each procedure revision builds one immutable environment. Domain-relevant,
case-insensitive names are assigned deterministic revision-local `SymbolID`
values from sorted canonical names. CFG blocks and edges are indexed once using
deterministic `BlockOrdinal` values derived from the graph's stable block order.
Module/project constants, declarations, interned strings, URL/path values, and
Array dimension shapes belong to this immutable environment and are not copied
into every block state.

Mutable flow state is procedure-local and owned by the solver invocation. A
state must not retain parser leases, source buffers, mutable environment maps,
or references to an obsolete analysis revision. Scratch states and changed-slot
buffers are reusable only until the procedure invocation ends.

### State representation

The solver provides compact indexed slots with two physical representations:

1. dense values plus a presence bitset for small or sufficiently dense domains;
2. sorted sparse entries for wide, sparse domains.

The hybrid selector is deterministic: it uses dense storage when the domain has
at most 64 symbols or the precomputed touched-symbol density is at least 25%,
and sparse storage otherwise. Representation benchmarks cover dense, sparse,
and hybrid choices; changing these constants requires new benchmark evidence.
Slot values are scalar IDs, flags, or bounded handles to immutable interned
values. The ordinary Array adapter follows this rule and keeps dimension
shapes in the value lattice. The initial HTTP adapter intentionally retains
the legacy `httpAnalysisState` nested maps in one compatibility slot while
its semantic-unit scalarization is developed separately; this exception is
why the Issue #713 allocation target remains open.

### Joins and convergence

Domain adapters implement transfer, lattice join, equality, and edge-specific
refinement. Joins update the destination state in place and return the changed
`SymbolID` values through a caller-owned buffer. A union-key map or complete
temporary result state is forbidden on the hot path.

The HTTP and Array lattices are finite and use their existing conservative
unknown/intersection/union semantics; no widening operation is added. The
solver converges when no destination slot changes. Idempotent joins report no
changed slots.

The worklist is indexed by `(BlockOrdinal, LaneOrdinal)` and ordered
deterministically. A lane/block is requeued only when its incoming state
changes. Exceptional and uncertain edges retain the predecessor input state
unless the domain adapter explicitly preserves an existing narrow exception.
Cancellation is checked at queue and long transfer boundaries. The solver does
not create goroutines or rule-level worker pools; it runs within the existing
procedure execution budget. The processed work-item trace is opt-in for
regression tests and is disabled by production adapters, so loop revisits do
not accumulate an unused history allocation.

### Evidence and projections

Source ranges, representative paths, credential sink lines, diagnostic
multiplicity, and other witness data are not lattice state. HTTP reconstructs
findings after convergence from stable states. Array keeps a separate witness
sidecar/replay and canonicalizes evidence by the existing source/range/order
rules. Semantic state may retain only bounded membership required to decide a
later transfer (for example an interned executable-path set); evidence itself
must not cause state copies or affect equality.

### Migration boundary

The first migration covers the HTTP scheduler boundary (`solveHTTPStates`)
and the ordinary block-level `arrayFlowState` walk. HTTP's nested state is
not yet compacted into semantic-unit slots. Issue #721 amends this boundary
for the remaining Array paths without changing the semantic lattice. The
source-line lifecycle lane, normal-edge refinement, exceptional/uncertain
edges, and combined runtime/evidence lanes use the same indexed solver when
their participant index and transfer contract can be built before execution.
Existing diagnostic IDs, severity, ranges, messages, suppression, JSON, LSP
projections, batch/realtime parity, and performance-counter meanings remain
unchanged. Generic `internal/vba/dataflow`, CFG storage, cross-run caching,
and unrelated semantic domains remain outside this decision.

### Amendment: advanced Array paths on the indexed solver (Issue #721)

The shared solver boundary now accepts source-line transfers and an explicit
edge-policy result. A policy may refine a normal edge, retain the predecessor
input for an exceptional or uncertain edge, or suppress propagation for an
edge that is not a continuation after a completed `Stop`/termination. A nil
policy retains the historical default of propagating the predecessor input;
adapters that do not need refinement therefore remain source-compatible. Edge
refinement is evaluated only for normal continuation edges. This keeps
conditional allocation, allocation guards, and module-configuration branches
out of exceptional and uncertain propagation.

The Array adapter indexes one fixed semantic participant catalog and transfers
through an internal cursor over compact slots. The legacy map adapter keeps
the same transfer, join, module-call effect, ByRef/return summary, and
evidence-read callback contract for differential tests, while retaining map
storage as the compatibility oracle. Joins update reusable output and edge
scratch in place. The compact solver boundary does not materialize a new
whole-state map for block scheduling; domain callbacks may retain their
documented copy-on-write behavior for legacy-compatible branch refinements.
Dimension and preserve shapes are immutable values; an update creates a new
shape value only when the shape itself changes.

The base and runtime lanes for one `CFGView` are solved as one indexed lane
group. The `VBA227` source-line lane that excludes the normal continuation is a
separate solver group so physical source-line order remains observable. A
completed stop suppresses all successor edges. Finding buffers, runtime
findings, ByRef-call evidence, module-entry contributions, and return
candidates remain lane-local sidecars rather than fixed-point state; the
existing deduplication and `sortFindings` stages remain authoritative for
evidence and result ordering.

Exceptional and uncertain edges retain the legacy conservative rule: they
receive predecessor input state. Under `On Error Resume Next`, the adapter may
retain predecessor output only when `ReliableExceptional` proves a plain
`ReDim`, `Array`, `Split`, or `Filter` allocation. `ReDim Preserve` does not
receive that exceptional allocation proof when its prior allocation or shape
is unknown. This rule applies equally to the source-line and combined lanes.

Compatibility selection is a preflight decision. `auto` selects the indexed
solver when the participant catalog is fixed, the cursor can represent every
semantic write, and the requested lane contracts are supported. Index-build
failure, an empty semantic state that still needs declaration-side diagnostics,
or an explicitly unsupported transfer contract selects the legacy walker.
Recovered or incomplete CFG input is not by itself a fallback reason when it
can be indexed and retains the same unknown/uncertain semantics. Cancellation,
transfer failure, or another execution-time error does not trigger a partial
legacy retry. The legacy walker remains available for compatibility and
differential testing; forced `compact` and `legacy` strategies are test and
benchmark controls only.

Compact/legacy walk counts and fallback reasons (empty-state, index-build, or
unsupported-contract) are developer-only telemetry.
They are not part of semantic state, normal CLI JSON, LSP payloads, corpus
snapshots, or the diagnostic review ledger. The specified ROneCOne benchmark
and alloc-space profile were collected against the #713 baseline. The focused
three-leaf profile shows a material reduction; end-to-end wall time remains
observational because host-load variance is high, and the profile paths and
values are retained in the corpus specification.

## Consequences

- Ordinary Array state copies become proportional to compact slot values or
  active sparse entries. HTTP gains deterministic indexed scheduling and
  cancellation through the shared boundary, but its compatibility slot still
  copies nested maps until the follow-up scalarization.
- The solver and domain adapters have a narrow ownership boundary that can be
  reused by later semantic kernels without exposing analyzer or protocol APIs.
- Deterministic indexed scheduling and changed-slot reporting make convergence
  and reprocessing measurable and reproducible.
- A second evidence projection is required where the old implementation
  collected findings during transfers; this preserves diagnostics while keeping
  witnesses out of the fixed-point state.
- The hybrid threshold and intern tables add a small setup cost. Dense, sparse,
  wide, loop-heavy, and real-world benchmarks must guard small-module
  regressions.
- Unsupported or recovered syntax continues to use the existing conservative
  edge and unknown-state behavior. No diagnostic precision is traded for the
  allocation reduction.

## Alternatives Considered

1. **Keep one map/deep-copy implementation per domain** — Rejected because the
   measured allocation cost is dominated by repeated state ownership and join
   work, and the two domains would continue to evolve different convergence
   contracts.
2. **Use dense arrays for every procedure** — Rejected because wide procedures
   with few domain symbols would pay for unused slots; the representation must
   be benchmarked for sparse and mixed shapes.
3. **Use a pointer-heavy persistent map or tree** — Rejected because it adds
   indirection and allocation pressure to the hot path without a demonstrated
   benefit, and it conflicts with the issue's compact-state goal.
4. **Put diagnostic evidence in the shared state** — Rejected because witness
   growth changes fixed-point equality, copies, and representative ordering.
5. **Add per-rule or nested solver workers** — Rejected because the existing
   bounded procedure budget and deterministic merge contract already define the
   concurrency boundary.
6. **Migrate generic source-to-sink dataflow or all semantic domains now** —
   Rejected because Issue #713 requires measurements for HTTP and Array
   coverage and a
   broader migration would expand compatibility risk without profile evidence.
7. **Persist solver results across runs** — Rejected because invalidation and
   compatibility are a separate cache decision.

## Evidence

- `internal/analyze/http_transport.go` (`httpAnalysisState`,
  `solveHTTPStates`, `cloneHTTPState`, `joinHTTPState`).
- `internal/analyze/array_safety.go` (`arrayFlowState`,
  `walkArrayCFGWorklist`, `walkArrayCFGCombined`, `cloneArrayState`,
  `meetArrayState`).
- `internal/analyze/http_transport_test.go`,
  `internal/analyze/array_participants_test.go`,
  `internal/analyze/array_kernel_benchmark_test.go`,
  `internal/analyze/procedure_planner_test.go`, and the existing lane profiling
  tests.
- `docs/specs/vba-analysis-ir.md` and
  `docs/specs/static-analysis-corpus.md` for ownership, deterministic analysis,
  and ROneCOne measurement contracts.
- ADR-0021, ADR-0022,
  `docs/adr/ADR-0044-bounded-procedure-effect-summaries.md`, and ADR-0046.

## Related

- Issue #713: `perf(dataflow): introduce a compact semantic state solver`
- Issue #693: shared semantic analysis kernel performance wave
- ADR-0021, ADR-0022,
  `docs/adr/ADR-0044-bounded-procedure-effect-summaries.md`, ADR-0046, and
  ADR-0048
