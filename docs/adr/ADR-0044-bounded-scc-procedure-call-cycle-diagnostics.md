# ADR-0044: Bounded SCC Procedure Call-Cycle Diagnostics

## Status

Accepted

## Context

ADR-0035 made `VBA244` exhaustive and deterministic by enumerating every
elementary directed cycle. That is useful for graph inspection, but the number
of simple cycles can be exponential in a dense graph. A normal `xlflow analyze`
command needs predictable latency even when a project has only a moderate
number of procedures and edges.

The existing diagnostic context remains valuable: users need a readable cycle
witness, and clients need structured path, edge, module-boundary, event,
effect, and uncertainty data. Changing the detector must not turn unresolved
or ambiguous calls into confirmed dependencies, or make ordinary recursion
look more dangerous than it is.

## Decision

Normal batch `VBA244` analysis uses strongly connected components (SCCs) over
the confirmed project-local call graph. SCC discovery is linear in vertices and
edges. A component is cyclic when it contains more than one procedure or a
self-edge. The analyzer emits exactly one finding per cyclic SCC before
suppression; it does not enumerate the component's elementary cycles.

Nodes and confirmed edges are canonicalized by stable procedure identity and
earliest call-site location. SCCs, members, and witnesses are sorted using the
same ordering and never depend on Go map iteration. Each SCC gets one closed,
deterministic representative witness: the canonical root is the least member,
the first stable outgoing edge is selected, and a deterministic breadth-first
path returns to the root. Witness construction is bounded by the SCC graph and
does not enumerate alternative cycles.

The `call_cycle` JSON shape remains additive and backward-compatible. Its
`path` and `edges` describe the representative closed witness; `cross_module`,
event-handler identities, dangerous effects, and resolution uncertainty are
deduplicated and aggregated from all SCC members and confirmed edges. Severity
preserves the existing contract: ordinary SCCs are `information`, while event
participation or an existing dangerous reachable effect elevates the finding to
non-blocking `warning`; uncertainty alone does not elevate it.

The existing exhaustive cycle enumerator remains available to explicit graph
inspection/debugging surfaces whose contract requires all elementary cycles.
The `impact` cycle payload is unchanged apart from its existing additive edge
data. Normal `analyze` and `check` must not invoke exhaustive enumeration.

## Consequences

- Analysis latency and memory are bounded by the call graph rather than by the
  number of elementary cycles, including dense cyclic SCCs.
- A project with multiple elementary cycles in one SCC receives one finding;
  this intentional multiplicity change is part of the `VBA244` scalability
  contract.
- The representative path is sufficient to explain recursion while aggregated
  SCC context preserves severity and machine-readable risk provenance.
- Consumers that require every possible path must use the explicit inspection
  surface; `VBA244` is no longer an exhaustive graph report.
- Deterministic ordering and cancellation remain required. A cancelled
  analysis returns no partial cycle findings.

## Alternatives Considered

1. **Continue exhaustive enumeration in normal analysis.** Rejected because
   elementary-cycle output can grow exponentially and gives no predictable
   latency bound.
2. **Report only one arbitrary back-edge or DFS cycle.** Rejected because map
   and traversal order would make findings unstable and could omit a useful
   return path.
3. **Use a fixed global cycle-count cap.** Rejected because it makes finding
   multiplicity depend on unrelated graph ordering and can truncate a witness
   without explaining which component remains covered.
4. **Drop structured effect and uncertainty context.** Rejected because
   severity and consumers rely on that provenance even when the path is
   representative rather than exhaustive.

## Evidence

- Issue #676 and parent issue #670 define the bounded-analysis requirement.
- `docs/specs/vba-procedure-call-cycles.md` defines the public `VBA244`
  contract and finding multiplicity.
- `internal/vba/callgraph` provides the confirmed project-local graph and keeps
  exhaustive enumeration for explicit inspection.
- Focused call-graph and analyzer tests cover self-recursion, independent and
  dense SCCs, deterministic witnesses, cancellation, and large sparse graphs.
- Developer benchmarks compare the pre-change exhaustive detector with the
  SCC detector using the same dense and 1000/2000-procedure fixtures.

## Supersedes

- `docs/adr/ADR-0035-procedure-call-cycle-analysis.md` (normal-analysis cycle
  enumeration and finding multiplicity)

## Superseded by

- None

## Related

- Issue #676
- Issue #670
- ADR-0024, ADR-0025
