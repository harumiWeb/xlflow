# VBA Procedure Call-Cycle Analysis

<!-- xlflow-rule-contract: {"id":"VBA244","family":"analyze","category":"reliability","default_severity":"information","scope":"project-wide","realtime":false,"configuration_key":"detect_procedure_call_cycles","inline_suppressible":true,"preflight_blocking":false} -->

`VBA244` reports confirmed recursive and cyclic dependencies between project
procedures. It runs in batch `analyze` and `check` and is enabled by default.
Disable it with `[analyze].disabled_rules = ["VBA244"]` or suppress an anchor
call with `xlflow:disable-line VBA244` / `xlflow:disable-next-line VBA244`.

## Graph contract

The detector reuses the project call graph. Only a uniquely resolved,
project-local call is a confirmed edge. Ambiguous, unresolved, external, and
dynamically bound calls never create or close a confirmed cycle. Normal batch
analysis computes strongly connected components (SCCs) over that graph. An SCC
is cyclic when it contains more than one procedure or a self-edge. This covers
self-recursion, mutual recursion, cycles crossing modules, dense/chorded
graphs, and independent SCCs without enumerating every elementary cycle.

Nodes and parallel endpoints are canonicalized by stable procedure identity and
earliest stable source location. SCC order, member order, edge order, and
finding order are deterministic and do not depend on Go map iteration. Each
cyclic SCC produces one deterministic representative closed witness. The
witness starts at the least stable procedure identity, follows the first stable
outgoing edge, and returns to the root through a deterministic path. Alternative
simple cycles in the same SCC are intentionally not reported by normal
`analyze`/`check`.

The explicit graph-inspection surface may retain exhaustive elementary-cycle
enumeration where its contract requires all paths. That inspection behavior is
separate from the bounded `VBA244` analyzer contract.

## Severity and context

An ordinary cycle is reported at `information`. It becomes a non-blocking
`warning` when a cycle procedure has an event-handler identity or a direct or
propagated reachable effect of Application-state mutation, error suppression,
workbook acquisition, or VBA file acquisition. Uncertainty is retained as
context but never by itself escalates severity.

The JSON finding has a `call_cycle` object containing the closed representative
procedure `path`, ordered `edges`, `cross_module`, event-handler identities,
dangerous effect evidence, and unresolved-call uncertainty. The finding
location is the first outgoing call in the representative witness; one finding
represents one cyclic SCC before suppression. Context fields are aggregated
from all SCC members and confirmed edges, so a dangerous effect or event
handler outside the selected witness still affects severity and remains
available to consumers. Uncertainty alone never elevates severity.

`path` nodes contain `qualified_name`, `module`, `kind`, `module_kind`, `file`,
and declaration `line`. The final node repeats the first node. Each `edges`
entry contains `caller`, `callee`, and the source `file`, `line`, `column`,
`end_line`, and `end_column`; edge `i` leaves path node `i`. `event_handlers`
contains stable participating identities. `dangerous_effects` contains the
effect `kind`, originating procedure, source file and line, and an optional
normalized target. `uncertainty` contains the resolution `kind`, origin,
unresolved callee text when available, and call location. Empty optional
collections are omitted, so adding this object remains backward compatible for
consumers that do not understand `VBA244`.

The impact command's existing `cycles[].nodes` remains an open path for
backward compatibility; `cycles[].edges` is additive and aligns with the
ordered nodes, while `VBA244` uses the closed `call_cycle.path` form.

The cycle graph is structural and is not an all-path execution proof. Effect
evidence remains conservative and follows the existing procedure-effect
summary contract. `opens_file` records reachable, non-recovered VBA
`Open ... For ... As #...` acquisition statements; `VBA219` continues to own
resource ownership and leak diagnostics.

## Scalability contract

SCC discovery is O(V + E). Representative witness construction is bounded by
the SCC's confirmed vertices and edges and never depends on the number of
elementary cycles. Consequently, the number of unsuppressed `VBA244` findings
is at most the number of cyclic SCCs, with exactly one finding per cyclic SCC
when no suppression applies. The analyzer must not invoke exhaustive cycle
enumeration on the normal `analyze` or `check` path.

The exhaustive detector can remain available behind an explicit inspection or
debugging surface. Performance coverage includes self recursion, a simple
cycle, independent SCCs, dense SCCs with many alternative simple cycles, large
acyclic graphs, and 1000–2000 procedure graphs. Benchmark results are retained
with the implementation change; timing thresholds are not part of CI.

The local Windows/amd64 benchmark used the same Go toolchain and fixture for
the old exhaustive detector and the bounded detector:

```text
go test ./internal/vba/callgraph -run '^$' -bench '^BenchmarkCycleDetection$' -benchmem -benchtime=1x -count=5
```

For the 8-node complete SCC, median detector time fell from 1.043 s,
736,920,616 B, and 24,534,409 allocations to 41.6 µs, 42,096 B, and 505
allocations (25,072× faster; 99.9960% less time). Bounded runs remained
linear-shaped for dense 100/250/1,000/2,000-node SCCs, 1,000/2,000-node
rings, and 1,000/2,000-node DAGs; the dense SCC result count was one. These
are local performance observations, not CI thresholds. The current benchmark
also reports 16,064 elementary cycles/op for the complete SCC baseline versus
1 component/op for the bounded detector; graph materialization is outside the
timed detector region.

For an end-to-end analyzer comparison, the existing single-module cyclic
fixtures were measured three times before and after the change with the same
filter and `-benchtime=1x`:

```text
go test ./internal/analyze -run '^$' -bench '^BenchmarkSingleModuleSynthetic/cyclic/(500|1000)-procedures$' -benchmem -benchtime=1x -count=3
```

The recorded pre-change medians were 999.321 ms/op for 500 procedures and
3,331.640 ms/op for 1,000 procedures. Post-change medians were 1,004.008
ms/op and 3,271.335 ms/op respectively: +0.47% and -1.81% relative change
within ordinary whole-analyzer variance. The detector-only benchmark above is
the meaningful scalability measurement because parsing and procedure-local
diagnostics dominate these end-to-end fixtures.

The change from one finding per elementary cycle to one finding per cyclic SCC
is intentional. Snapshot and corpus reviews must treat disappeared elementary
cycle anchors as multiplicity consolidation, while any new finding must still
be checked for a confirmed source-local edge and false-positive regressions.
