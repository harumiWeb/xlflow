# ADR-0035: Deterministic Procedure Call-Cycle Analysis

## Status

Accepted

## Context

The existing call-graph impact command exposes only target-scoped DFS
back-edges. It omits independent project cycles, can miss distinct nested
paths, and represents cycles as an unordered node set. Reliability analysis
needs complete recursive paths, deterministic output, and enough effect and
uncertainty context to distinguish ordinary recursion from dangerous cycles.

## Decision

Generalize `internal/vba/callgraph` with a graph-wide, cancellation-aware
elementary-cycle enumerator. It uses only uniquely resolved project-local
edges, deduplicates parallel endpoints by earliest source location, preserves
directed path order, and canonicalizes rotations at the lowest stable node.
The existing impact API keeps its shape while receiving ordered cycle nodes and
additive edge data.

Add default-enabled, batch-only `VBA244`. Emit one finding per canonical cycle
at its first outgoing call. Ordinary cycles use `information`; event-handler
participation, Application-state mutation, error suppression, workbook
acquisition, or VBA file acquisition in direct or propagated reachable effect
evidence elevates the finding to a non-blocking `warning`. Unresolved and
ambiguous calls remain uncertainty and cannot prove or elevate a cycle.

Expose a structured `call_cycle` finding context so clients can consume the
closed path, call locations, module boundary, effects, and uncertainty without
parsing diagnostic prose. File acquisition is added to the shared effect
vocabulary while `VBA219` retains ownership/lifetime analysis.

## Consequences

- Recursive, nested, cross-module, and independent cycles are visible with
  stable paths and no rotation duplicates.
- Dense graphs can contain many elementary cycles; the implementation supports
  cancellation and does not silently truncate the result.
- The analyzer JSON contract gains an additive `call_cycle` field and the
  impact cycle payload gains additive edge data.
- Dynamic or unresolved dispatch remains conservative and may leave cycle
  uncertainty without producing a confirmed dependency.
- LSP publication and compile/VBE validation are outside this decision; the
  rule is batch-only and non-blocking.

## Alternatives considered

1. Report one representative cycle per SCC. Rejected because nested/chorded
   directed paths would be lost.
2. Keep the existing DFS back-edge implementation. Rejected because it is
   target-scoped and can omit distinct elementary cycles.
3. Treat unresolved calls as cycle edges. Rejected because uncertainty is not
   proof of a dependency.
4. Encode the path only in message prose. Rejected because consumers need
   stable machine-readable path, effect, and uncertainty data.

## Evidence

- `internal/vba/callgraph/callgraph.go` and its deterministic cycle tests.
- `internal/vba/effects` summaries and `docs/specs/vba-effect-analysis.md`.
- `internal/analyze/procedure_cycles.go` and focused `VBA244` tests.
- `docs/specs/vba-procedure-call-cycles.md`.

## Related

- Issue #459
- ADR-0021, ADR-0023, ADR-0024, ADR-0025-rooted-vba-callgraph-reachability
