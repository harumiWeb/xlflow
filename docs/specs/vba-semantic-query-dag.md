# VBA Semantic Query DAG

This specification defines the internal revision-scoped semantic query store
for Issue #715. ADR-0048 records the rationale. The store reuses immutable
semantic kernel and diagnostic-projection results within one process while
preserving the ownership, conservative uncertainty, cancellation, and public
output contracts of the existing analyzer.

## Scope and boundaries

The query store may be shared by batch and LSP analysis when they refer to the
same immutable project/revision inputs. It is an execution optimization, not a
new analyzer API. It does not change diagnostic IDs, severity, ranges, messages,
evidence, multiplicity, suppression, ordering, exit status, CLI options,
configuration keys, JSON envelopes, or LSP schemas.

The store is process-local. It does not persist values to disk, define a cache
file or serialization format, or share entries across processes. Parser and
tree-sitter reuse, cold parsing optimization, and distributed caching remain
outside this specification.

## Ownership and revision lifetime

The store owns immutable Go values and dependency metadata. Query values may
refer to immutable IR/CFG and capability data only through an owner that keeps
that revision alive; they must not retain parser leases, tree-sitter nodes,
source buffers, mutable maps or slices, solver scratch, finding builders, or
obsolete revision objects. Mutable worklists and evidence reconstruction stay
inside one evaluation and are never placed in a reusable value.

Each evaluation receives an immutable revision handle covering the document or
project inputs, module facts, resolution state, capabilities, relevant
configuration, and cancellation context. A successor revision may inherit a
completed value only when its key and every recorded dependency fingerprint
match. Close/reopen begins a new lifecycle and cannot inherit lifecycle-local
artifacts. Eviction must retain values reachable by the current or actively
referenced revisions and may release unreachable older entries.

## Query keys and dependency graph

A query key contains these components:

1. canonical procedure identity: normalized document/module identity,
   qualified name, procedure kind, and deterministic declaration identity;
2. procedure fingerprint: body, signature, and module/conditional context
   components, plus any resolution/call-edge fingerprint used by the query;
3. semantic query identity: kernel or diagnostic projection, including its
   versioned internal meaning;
4. relevant configuration hash; and
5. relevant project capability/resolution revision.

The exact internal encoding is not public, but omitting a semantic input from
the key or dependency list is a correctness defect. Dependency edges identify
the immutable source/module facts, capabilities, procedures, call edges,
effect summaries, kernels, and projections read during evaluation. Revision-
local IR/CFG IDs aid lookup only; they are not persistent workspace identities.

The graph has source/module and project-capability leaves, procedure semantic
kernels, and diagnostic projections. A projection reads the immutable result
of its required kernel and does not silently construct a duplicate semantic
result. One key has at most one in-flight build; concurrent readers share that
build and receive the same immutable value.

## Invalidation and red/green reuse

When a new revision is opened, changed leaves and reverse dependents are
identified from the union of the previous and current dependency edges. The
union is required so removal, redirection, or resolution of a call invalidates
the former caller as well as a new target. Candidates are marked red and
reevaluated in deterministic dependency order.

After reevaluation, an unchanged output fingerprint turns a red entry green
and stops invalidity propagation through that entry. A changed output keeps
its reverse dependents invalid and schedules them for recomputation. A hit is
valid only after key, dependency, ownership, and revision-lifetime checks.

The minimum invalidation matrix is:

| change                                                   | required invalidation boundary                                  |
| -------------------------------------------------------- | --------------------------------------------------------------- |
| procedure body or local source range                     | that procedure's affected kernels and projections               |
| signature or declaration identity                        | the procedure and callers/callees that consume the signature    |
| resolution, ambiguity, or call addition/removal/redirect | affected procedures, reverse callers, and dependent summaries   |
| effect or project-summary output                         | consumers recorded on the changed summary edge                  |
| module declarations or conditional context               | affected module and dependent resolution/CFG/capability queries |
| relevant configuration                                   | queries whose key includes that configuration input             |
| recovered, dynamic, incomplete, or unknown ownership     | smallest reliable module, SCC, or project boundary, fail-open   |

The table describes dependency boundaries, not a promise that every uncertain
case can be reduced to a procedure. A safe implementation may invalidate a
larger boundary when it cannot prove a narrower one; it must not reuse across
an unproven boundary.

## Cancellation and concurrent revisions

Evaluations check cancellation at query scheduling and long-running semantic
work boundaries. A canceled, failed, or panicked build commits neither its
value nor its dependency edges. Waiters may retry the key, and no partial
diagnostic projection is publishable.

Revision publication is generation-safe: a late result from revision N cannot
replace, publish over, or become the completed value for newer revision N+1.
Same-revision readers may run concurrently against immutable values; mutable
solver state, findings, telemetry deltas, and projection buffers remain local
to the evaluation. The query store does not add nested rule-level workers or
expand the existing bounded execution budget.

## Diagnostics and telemetry

The cache-enabled result must be structurally equivalent to a fresh full
recomputation, including diagnostic ordering, evidence, suppression, and
batch/LSP projection. Finding evidence is rebuilt or defensively copied at
the existing publication boundary; it is not shared as mutable query state.

The following additive counters use the existing stderr/performance-log
channel:

- `semantic_query_hits`;
- `semantic_query_misses`;
- `semantic_query_invalidated_procedures`; and
- `semantic_query_recomputed_kernels`.

They are developer-only telemetry. They never appear in normal CLI JSON, LSP
diagnostic payloads, corpus snapshots, or `reviews/diagnostics.jsonl`, and they
must not affect findings or snapshot identity. Repeated runs should report
deterministically ordered records for a fixed revision and input.

## Verification requirements

Focused tests must cover key stability, dependency recording, reverse-edge
traversal, output-fingerprint green recovery, single-flight evaluation,
retryable cancellation, revision eviction, and late-result rejection. The
invalidation matrix must cover body, signature, configuration, resolution,
callee/effect summary, call add/remove/redirect, and module/project changes.

Differential tests compare warm incremental and LSP/batch results with a fresh
full recomputation, including findings, ranges, evidence, multiplicity,
ordering, suppression, exit status, and normal JSON/protocol projection.
Concurrent-revision tests and `go test -race` must cover cancellation and
replacement. Benchmarks record cold/full, warm/same-revision, local body,
signature, resolution, configuration, and many-file edits; wall-clock values
are observations rather than CI thresholds. Corpus verification is read-only:
any unexplained snapshot or review-ledger delta is a stop-and-investigate
condition, and query counters are never copied into corpus artifacts.

## Related

- `docs/adr/ADR-0048-revision-scoped-semantic-query-dag.md`
- `docs/specs/vba-analysis-ir.md`
- `docs/specs/vba-control-flow-graph.md`
- `docs/specs/cli-contract.md`
- `docs/specs/static-analysis-corpus.md`
- ADR-0021, ADR-0022, ADR-0046, and ADR-0047
- Issue #715
