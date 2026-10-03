# Source-Only Project Architecture Report

## Command and ownership

```text
xlflow architecture
xlflow architecture --json
xlflow architecture --module Billing
xlflow architecture --path src/modules --json
```

`architecture` collects one informational source snapshot. It reads configured
`[src]` roots and the conventional `tests/` tree, using the same UserForm
embedded/sidecar selection as metrics. It does not open a workbook, invoke the
bridge or COM, execute VBA, generate TypeLib metadata, or write source/project
state. Missing optional external metadata preserves conservative resolution.

The report is a protocol-neutral projection of canonical symbols,
DocumentIR/CFG, resolution, dependencies, rooted reachability, cyclic SCCs,
procedure metrics, hotspots, mutable-state facts, and effect summaries.
Neither CLI rendering nor architecture collection owns a second analyzer.
The command never invokes the full lint/analyze pipelines or emits their
findings, MX001, or MX002. Configuration remains validated normally.

`[metrics].exclude`, metric thresholds, hotspot selectors, and diagnostic
enablement/suppression do not select architecture facts or determine its exit
status. Metrics-only exclusions can make `metrics` use a different cohort;
the shared counting and score definitions remain identical for equal inputs.

## Schema and evidence

The common JSON envelope contains `command="architecture"` and an
`architecture` object with `schema_version=1`. The report contains:

- project and selected-scope summaries;
- modules and procedures with canonical identities, declaration ranges,
  visibility, metrics, and module/procedure kinds;
- entry-point evidence with confirmed/possible confidence and reason;
- confirmed typed dependencies and uncertain dependency evidence;
- cyclic SCCs with every member and one deterministic cycle witness;
- definitely unreachable internal procedures and possible reachability;
- the native hotspot report, retaining its schema version and score model;
- reusable mutable-state field/access facts and direct Excel effect evidence;
- external reference evidence and call uncertainty counts.

Module/procedure IDs use the existing dependency projection. Hotspot IDs retain
their own canonical metrics representation; consumers join by file,
declaration position, and procedure kind instead of comparing unrelated ID
strings. Procedure kind distinguishes Property Get, Let, and Set. Declaration
locations retain canonical one-based lines and columns; procedures also retain
the original declaration range, including byte offsets.

Adding fields is compatible with schema version 1. Removing/redefining fields
or changing identity meaning requires a new architecture schema version.
Nested hotspot schema/model versions retain their independent ownership.
Empty collection fields serialize as `[]`; source paths are project-relative
with `/` separators. Telemetry and timings do not appear in stable report JSON.

### Counts and confidence

`confirmed_call_edges` counts distinct confirmed procedure-to-procedure
relationships. Repeated source call sites remain in dependency evidence;
aggregated module edges are not counted again. Typed dependencies retain
`calls`, `uses_type`, `constructs`, and `implements` categories.

Uncertain calls retain their original resolution status and locations.
External, builtin-like, member, unresolved, ambiguous, and dynamic relationships
never become invented confirmed project edges. Dynamic callback evidence keeps
the API, expression, known target where available, and source range.
Callback extraction uses the resolved snapshot so receiverless and
`With Application` `Run`/`OnTime`/`OnKey` calls retain their targets. Regression
controls also protect project-procedure and non-callable local shadows from
being interpreted as Application callbacks.

`dependency_cycles` counts cyclic SCCs, including direct recursion. The
representative witness is one deterministic cycle, not every elementary cycle.
Hotspot `cycle_count` keeps its existing bounded elementary-cycle semantics
and may be a conservative lower bound when its work budget is exhausted.

Roots and unreachable classification follow the canonical reachability model,
including public API possibilities, host events, test hooks, exposed classes,
and static/unknown callback references. An event-like name alone is not new
evidence. A configured ambiguous root remains possible, and a missing root is
retained as uncertainty rather than labeled a confirmed resolved procedure.

The hotspot `excel_effect_count` is the existing syntax-based reference signal.
It is distinct from direct side-effect evidence collected by the canonical
effects model. Effects and mutable-state details preserve the collector's
supported semantics; they do not imply that code ran or that a defect exists.
Project/module effect totals count origin evidence once, without inflating
counts by materializing the same propagated effect at every caller.

Module-state summary counts are distinct procedure/field access relationships,
and each field retains reader, writer, and collection-mutator identities.
Constants remain in the field inventory with their declaration kind. Native
hotspot mutable-state signals retain their per-access occurrence counting model;
they are not replaced with the field summary counts.

## Display scope

`--module` matches an exact module name case-insensitively. `--path` selects a
source file or directory; relative paths resolve from the project root. When
both are supplied, selected entities must satisfy both filters. Invalid
module names, nonexistent paths, and paths outside the project are input
errors. A valid empty selection is an empty report view.

All sources are parsed and resolved before filtering. Project-wide reachability,
fan-in/fan-out, metrics, and hotspot ranks/scores survive filtering unchanged.
`project_summary` describes the whole analyzed project; `summary` describes the
selected view. The scope metadata identifies the display filters.

Confirmed dependencies incident to selected entities retain their endpoint
nodes. `dependency_boundary_node_ids` identifies out-of-scope endpoint nodes. Related
cyclic components retain complete member/witness information. Boundary context
does not increase the selected module/procedure counts. Filtering never turns
an out-of-scope target into an unresolved call or proves a procedure unreachable.

## Human output and failure behavior

Human output starts with a compact summary and shows entry points, dependency
counts, procedure/module hotspots, unreachable procedures, state/effect counts,
and uncertainty. Each large entity list shows at most ten entries in the same
order as JSON, with a remaining count. JSON keeps the complete selected report.

Successful reporting exits 0 even with cycles, unreachable code, uncertainty,
mutable state, or high hotspot scores. Invalid configuration/scope exits 2;
parse recovery exits 1; read/operation/cancellation failures exit 3. Failures
use the standard failed envelope and omit a partial architecture report.

## Construction and verification contract

Code source documents are read and parsed once per command revision. Parsed
state supplies canonical symbols and IR; resolver/index/CFG state is shared
between projections. Designer metadata is read separately only when needed.
Source snapshots describe captured input, not an atomic filesystem transaction.

Graph assembly reuses confirmed relationships. SCC condensation and shared
module closure prevent a separate full graph traversal per procedure for
affected-module counts. Cycle enumeration retains its deterministic bounded
policy. Effects use direct evidence APIs and do not eagerly reconstruct every
procedure's propagated provenance. Cancellation is checked in discovery,
parsing, resolution, and aggregation.

Tests protect equality with existing projections for identical source sets,
all module kinds and root confidence classes, uncertainty, effects/state,
filter boundaries, threshold independence, empty/failed inputs, serialization
order, repeated-run JSON equivalence, and parse/IR/CFG construction counts.
Synthetic sparse, chain, and dense-cycle projects cover scaling without flaky
wall-clock limits. Analyzer performance is compared against `origin/main` on
ROneCOne and std-vba, while architecture has its own end-to-end benchmark.

## Related

- ADR-0062
- `vba-call-graph-reachability.md`
- `vba-procedure-complexity-metrics.md`
- `vba-procedure-and-module-hotspots.md`
- `vba-module-state-analysis.md`
- `vitepress/commands/architecture.md`
