# ADR-0048: Revision-Scoped Semantic Query DAG

## Status

Accepted

## Context

The analyzer now has immutable procedure IR and CFG views, project capability
planning, semantic execution plans, and compact solver state. Those boundaries
share work within one analysis revision, but a subsequent edit still rebuilds
every affected kernel even when its semantic inputs are unchanged. Batch and
LSP also have separate entry-point lifecycles, so they need one explicit
ownership and invalidation contract before they can safely reuse results.

Issue #715 introduces reuse across immutable revisions in one process. The
reuse must follow semantic dependencies rather than file-level guesses. A body
edit should not invalidate unrelated procedures; a signature, resolution,
call/effect, module, or configuration change must invalidate every dependent
query whose input can change. Reuse must not weaken the conservative treatment
of recovered, ambiguous, dynamic, or incomplete VBA.

## Decision

Add a process-local semantic query store represented as a directed acyclic
dependency graph. The store is an internal execution facility that may be
shared by batch and LSP analysis in the same process. It does not become a
public CLI, JSON, or LSP interface and it does not define a disk cache format.

### Ownership and identity

The store owns only immutable, Go-managed query values and the dependency
metadata that led to those values. It must not retain tree-sitter nodes,
parser leases, source buffers, mutable solver scratch, finding builders, or
revision-local objects whose owner has retired. A revision handle identifies
the immutable IR/CFG, module facts, project capabilities, configuration, and
resolution view used for one query evaluation; callers retain ownership of
those inputs.

Query entries are scoped to a canonical procedure identity and a semantic
query. The identity includes the normalized document/module identity,
qualified procedure name and kind, and the procedure's deterministic
declaration identity. A query key includes, at minimum:

- the procedure fingerprint, with body, signature, and module/conditional
  context components kept distinguishable;
- the semantic kernel or diagnostic projection identity;
- a hash of configuration that can affect that query; and
- the relevant project capability/resolution revision.

Dependency edges record the exact source, module, capability, procedure, call,
effect-summary, kernel, and projection inputs read by a query. Procedure
fingerprints and revision-local IR/CFG IDs are not workspace-global identities;
the canonical identity and current dependency fingerprints are required before
an entry can be reused.

### Amendment: fingerprint and invalidation overhead reduction (#724)

The revision-owned analyzer projection prepares immutable configuration, TypeDB,
module, capability, procedure, effect, and call-resolution leaves once before
procedure workers run. The three reusable procedure lanes share those inputs;
they do not serialize the same large maps or plans again for each lane. A
procedure fingerprint contains procedure-local source and signature facts,
while module, capability, data-flow, effect, and resolution leaves remain
explicit dependencies at the kernel boundary. This separation may reduce a
dependency set only when the kernel does not read the removed leaf; it must not
weaken fail-open behavior for incomplete or ambiguous input.

The store's internal identity is a comparable `Key` rather than a serialized
string. Dependencies are canonicalized and deduplicated at publication, and
reverse edges plus pending/epoch/eviction metadata use the same key. Exact
document and procedure indexes service LSP invalidation, so a normalized path
does not require scanning every cached key. These are private representations;
the process-local lifetime, cancellation/single-flight behavior, late-result
rejection, diagnostic projection, and public output contracts remain those
decided by this ADR. No disk cache, cross-process format, or public API is
introduced.

### Query graph and invalidation

The graph contains immutable source/module and capability inputs, procedure
semantic kernels, and diagnostic projections. A kernel result is materialized
once for a matching key; projections consume the matching immutable result and
must not rebuild the kernel privately. Same-revision concurrent readers share
one in-flight evaluation for a key.

On a new revision, changed leaves and their reverse dependents are marked red
using both the previous and current dependency edges. This covers deleted or
redirected calls as well as newly added edges. Red entries are reevaluated in
deterministic dependency order. If an entry's output fingerprint is unchanged,
it becomes green and invalidity does not continue downstream; otherwise its
reverse dependents remain candidates for recomputation.

Invalidation is dependency-scoped: a body edit starts at that procedure and
its projections; a changed signature, resolution result, call edge, effect
summary, module context, capability, or relevant configuration reaches the
callers, callees, module summaries, and project summaries that actually depend
on it. Recovered syntax, ambiguous or dynamic binding, incomplete resolution,
and unknown ownership fail open at the smallest reliable module, SCC, or
project boundary. The DAG may therefore recompute more work when precision is
not proven, but it must never reuse a result across an unproven semantic
boundary.

### Cancellation, concurrency, and lifetime

Every evaluation observes its revision context. A canceled, failed, or
panicked evaluation publishes no query value and commits no new dependency
edges; waiters may retry the key. A result from an older revision cannot replace
or publish over a newer revision, even when its worker completes later. Values
inherited by a successor revision remain immutable and are released when no
revision or in-flight reader can reach them. Eviction is bounded to current and
actively referenced revision state; close/reopen starts a new lifecycle.

The store does not add nested rule workers or change the existing bounded
execution budget. Finding assembly, suppression, ordering, and protocol
publication remain outside the reusable semantic value and retain their
existing cancellation and defensive-copy boundaries.

### Compatibility and observability

Reuse is an execution optimization. A cache-enabled run must produce the same
diagnostic code, severity, range, message, reason, suggestion, evidence,
multiplicity, ordering, suppression result, exit status, and JSON/LSP projection
as a fresh full recomputation. No partial or stale finding may be published.

The performance recorder adds the developer-only counters
`semantic_query_hits`, `semantic_query_misses`,
`semantic_query_invalidated_procedures`, and
`semantic_query_recomputed_kernels`. They are emitted only through the
existing stderr/performance-log channel and never appear in normal CLI JSON,
LSP payloads, corpus snapshots, or the diagnostic review ledger.

The store is process-local and revision-scoped. Persistent disk caching,
cross-process reuse, cache files, serialization, and compatibility guarantees
for stored query values are explicitly out of scope.

## Consequences

- Unchanged procedures can reuse semantic kernels and projections after an
  edit, while dependency changes still propagate to the necessary callers and
  summaries.
- The store and reverse-edge bookkeeping add memory and graph-maintenance
  work, and conservative uncertainty may intentionally reduce reuse.
- Query keys and dependency recording become internal compatibility contracts;
  any new kernel or projection must declare its relevant inputs.
- Batch and LSP can share the same process-local policy, but independent
  processes and close/reopen lifecycles do not share values.
- Diagnostic evidence remains a projection concern, so reuse cannot alter
  finding identity or representative ordering.

## Alternatives Considered

1. **Rebuild every query for every revision** - Rejected because unchanged
   procedures pay the same kernel cost after local edits.
2. **Invalidate an entire file or project for every edit** - Rejected because
   it cannot meet the issue's local-edit acceptance criterion and hides real
   dependency boundaries.
3. **Keep one cache per diagnostic rule** - Rejected because kernels and
   capabilities are shared semantic inputs; per-rule caches duplicate work and
   can diverge in invalidation behavior.
4. **Persist or share query values across processes** - Rejected because it
   would require a disk schema, compatibility, corruption, and migration
   policy outside this issue.
5. **Reuse red entries without output comparison** - Rejected because a
   dependency may be invalidated conservatively even when its resulting value
   remains identical; red/green comparison preserves both safety and reuse.

## Evidence

- Issue #715: `perf(analyze): add a revision-scoped semantic query DAG`.
- Procedure ownership and revision-local identity:
  `docs/adr/ADR-0021-procedure-analysis-ir.md` and
  `docs/specs/vba-analysis-ir.md`.
- Conservative graph identity and immutable views:
  `docs/adr/ADR-0022-conservative-vba-control-flow-graph.md`.
- Capability, kernel, projection, and bounded execution contracts:
  `docs/adr/ADR-0046-procedure-applicability-planning.md` and
  `docs/adr/ADR-0047-compact-semantic-state-solver.md`.
- CLI telemetry and corpus verification contracts:
  `docs/specs/cli-contract.md` and
  `docs/specs/static-analysis-corpus.md`.

## Related

- Issue #715
- Issue #701 (semantic execution plans)
- Issue #713 (compact semantic state solver)
- Issue #714 (immutable CFG views)
- ADR-0021, ADR-0022, ADR-0046, and ADR-0047
- `docs/specs/vba-semantic-query-dag.md`
