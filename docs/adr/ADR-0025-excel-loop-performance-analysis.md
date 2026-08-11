# ADR-0025: Excel Loop Performance Analysis

## Status

Accepted

## Context

VBA code that reads or writes one Excel cell at a time can be dominated by
Excel COM round trips rather than by the VBA computation itself. The pattern is
easy to write with `Cells`, `Range`, `Offset`, `For Each cell In range.Cells`,
worksheet functions, or per-cell formatting, but a generic text search cannot
tell whether the receiver is an Excel object, whether a statement is inside a
reachable loop, or whether a range operation is already a bulk transfer.

Issue #452 requires a diagnostic that covers indexed and `For Each` loops,
distinguishes reads, writes, formulas, and formatting, retains nested-loop
depth as explanatory context, explains the COM cost, suggests array or
bulk-range operations, and avoids clearly trivial fixed-size loops. The rule
must also account for uniquely resolved local helpers without turning
ambiguous or late-bound VBA into false certainty.

Issue #453 adds a related but distinct cost: a loop can repeatedly resolve the
same workbook, worksheet, table, named-range, pivot, or chart member chain
even when the chain does not depend on the loop iterator. Those invariant
lookups should be hoisted into a cached local while iterator-dependent cell
access remains inside the loop.

Issue #454 adds the array-allocation counterpart: `ReDim Preserve` copies the
existing contents on every resize, so placing it in a reachable loop can turn
an otherwise linear fill into quadratic work. The diagnostic must cover every
supported VBA loop form, distinguish loop-variable growth from repeated
constant-size reallocations, account for nested-loop cost, and preserve the
existing `ReDim Preserve` and dimension-safety ownership boundaries.

Issue #455 adds a fourth cost shape: a costly operation may target an entire
row, column, worksheet, or an unbounded `UsedRange` when a bounded range was
intended. One-time whole-sheet formatting remains legitimate, so this rule
must be opt-in and must distinguish loop-contained processing from a one-time
operation.

## Decision

Add `VBA225`, a default-enabled `analyze` rule in the performance category. It
runs in batch analysis and the shared real-time editor path. It is a medium-
precision, interprocedural, non-blocking warning with normal inline
suppression. The rule uses the normalized procedure IR and conservative CFG
from ADR-0021 and ADR-0022. It does not parse source a second time, open Excel,
or add a new JSON field.

The rule recognizes reachable `For`, `For Each`, `While`, and `Do` loop bodies.
It resolves Excel receivers and members where the existing analysis context can
prove them, then classifies repeated loop-contained work as cell/range reads,
writes, formulas, formatting, address or worksheet lookups, and recognized
worksheet-function calls. A uniquely resolved project-local helper contributes
only when its direct or propagated summary contains one of those classified
accesses. Unresolved, ambiguous, external, and dynamic calls are uncertainty,
not positive evidence.

Emit one deterministic warning per loop at its loop header. Nested-loop depth
remains in the message and reason as context, but does not escalate severity or
make `VBA225` source-preflight-blocking.
The message and reason name the access kind, identify nested depth when
relevant, and explain the cost of a COM boundary per iteration. Suggestions
prefer one-time range transfers into VBA arrays, one range assignment for
writes or formulas, and range-level formatting.

Do not report a range-level transfer or formatting operation whose target and
shape are independent of the loop item. Also do not report a loop whose maximum
iteration count is statically provable as three or fewer iterations from
integer literals or simple constant integer arithmetic, or when a literal
`For Each` range contains at most three cells. These exemptions keep the rule focused on repeated cell-by-cell work
while leaving unknown bounds and dynamic object identity conservative.

Configure the rule with `[analyze].disabled_rules = ["VBA225"]`; preserve the
legacy `detect_excel_cell_access_in_loops` Boolean during the compatibility
window.
The rule is inline suppressible, does not block `push` or `run` preflight, and
uses the existing analyzer finding envelope. Its static-analysis registry entry
and generated diagnostic catalog are follow-up projections of this contract;
the generated catalog must be regenerated rather than edited manually.

Add `VBA238` for issue #453. It is a default-enabled, warning-level,
procedure-local performance rule in batch analysis and the shared real-time
path. It normalizes repeated Excel member chains, ignores chains that reference
the active loop variable or are only trivial local-variable access, and reports
an invariant chain that can be extracted into a cached local before the loop.
Nested loops are handled by requiring invariance at the selected loop boundary.
`VBA238` is non-blocking, inline suppressible, and configured with
`[analyze].disabled_rules = ["VBA238"]` (with the legacy
`detect_loop_invariant_excel_object_resolution` Boolean retained during the
compatibility window). `VBA238` owns lookup-hoisting guidance; `VBA225` retains
ownership of repeated per-cell or per-item object-model work.

Add `VBA241`, a default-enabled, procedure-local, warning-level performance
rule in batch analysis and the shared real-time path. It reuses the normalized
procedure IR, conservative CFG loop regions, and existing `ReDim Preserve`
dimension analysis to report reachable loop-contained resizes. A dimension
expression is loop-variable dependent when its IR variable accesses match the
current or any outer loop control variable, including accesses inside helper
expressions; helper procedure bodies are not followed transitively. A single
non-nested loop with loop-invariant bounds is reported at `information`, while
loop-variable growth or nesting depth two or greater is reported at `warning`.
The rule emits one finding at the resize statement, remains non-blocking and
inline suppressible, and is configured with
`[analyze].disabled_rules = ["VBA241"]` (with the compatibility key
`detect_redim_preserve_in_loops` retained during the compatibility window).
`VBA208` continues to own non-final-dimension correctness findings, and
`VBA227` continues to own allocation/lifecycle safety; neither rule is merged
with this performance diagnostic.

Add `VBA242` for issue #455. It is an opt-in, information-level by default,
procedure-local performance rule in batch analysis and the shared real-time
path. The detector recognizes full-row, full-column, full-sheet, and direct or
unbounded `UsedRange` shapes only when they feed an expensive operation such as
formula/value assignment, calculation, formatting, find/replace, or sorting.
Explicitly bounded ranges, bounded `Resize`, and bounded `Intersect` forms are
accepted. A reachable loop escalates the finding from `information` to
`warning`; the rule remains non-blocking and inline suppressible. Configure it
with `detect_expensive_full_range_operations = true` during the compatibility
window, while `[analyze].disabled_rules = ["VBA242"]` is the stable policy.
`VBA201`, `VBA217`, and `VBA225` retain ownership of find-result safety,
last-row calculation guidance, and repeated per-cell work respectively.

## Consequences

- Positive: common cell-by-cell loop patterns receive actionable guidance before
  Excel execution, without requiring a workbook or runtime sampling.
- Positive: CFG reachability, receiver resolution, and helper summaries reduce
  false positives from dead code, bulk range transfers, and unrelated VBA calls.
- Positive: nested-loop context makes the most expensive shapes visible while
  keeping the rule advisory and capped at warning severity.
- Negative: dynamic dispatch, late binding, unresolved aliases, and helpers that
  cannot be uniquely resolved remain unreported even when they are slow.
- Negative: interprocedural summaries add analysis work and must retain stable
  provenance so one loop does not produce duplicate helper findings.
- Negative: the fixed small-loop threshold is a deliberate heuristic; projects
  with unusual workloads may still need a local suppression.
- Negative: invariant-chain normalization is conservative around dynamic
  dispatch and ambiguous `With` receivers, so some cacheable lookups remain
  unreported.
- Positive: `VBA241` makes the quadratic-copy shape visible while reusing the
  established parser and CFG facts, and its separate severity classification
  distinguishes a loop-invariant resize from growth tied to an iterator.
- Negative: unknown or dynamically computed bounds can remain unclassified, and
  the rule cannot prove helper-procedure internals without violating its
  procedure-local scope. Projects may still need a local suppression for
  intentionally bounded reallocations.
- Positive: `VBA242` makes accidental worksheet-scale processing visible while
  preserving legitimate one-time formatting through opt-in configuration and
  explicit bounded-range exemptions.
- Negative: shape inference is intentionally conservative; dynamic addresses,
  named ranges, and aliases whose effective bounds are unknown remain
  unreported.

## Alternatives Considered

1. **Search source text for `.Value` or `Cells`.** Rejected because it cannot
   establish loop reachability, Excel receiver identity, read/write/formatting
   semantics, or bulk-range exemptions.
2. **Flag every Excel call inside every loop.** Rejected because cached roots,
   invariant bulk operations, small fixed loops, and non-cell Excel work would
   create noisy findings.
3. **Use runtime timing or Excel telemetry.** Rejected because `analyze` is a
   source-only command and runtime measurements are environment- and data-
   dependent.
4. **Make the finding block `push` and `run`.** Rejected because performance
   depends on workbook size and execution context; correctness and modal-error
   rules own preflight blocking.
5. **Analyze only direct statements and ignore helpers.** Rejected because a
   common refactoring moves the per-cell Excel call into a project-local helper,
   while the caller still pays for it on every iteration.
6. **Treat every `ReDim Preserve` as a correctness error.** Rejected because
   the operation is valid VBA; only a reachable loop establishes the repeated
   copy cost, and dimension correctness remains owned by `VBA208`.

## Evidence

- Requirements: xlflow issue #452.
- Related requirements: xlflow issue #453 (loop-invariant Excel object
  resolution and cache-hoisting guidance).
- Related requirements: xlflow issue #454 (loop-contained `ReDim Preserve`
  performance analysis).
- Related requirements: xlflow issue #455 (expensive full-row, full-column, and
  full-sheet operations).
- Analyzer ownership and public diagnostic policy:
  `docs/adr/ADR-0013-analyze-runtime-risk-ownership.md`.
- Procedure syntax and resolution: `docs/specs/vba-analysis-ir.md` and
  `docs/adr/ADR-0021-procedure-analysis-ir.md`.
- Reachability and loop semantics: `docs/specs/vba-control-flow-graph.md` and
  `docs/adr/ADR-0022-conservative-vba-control-flow-graph.md`.
- Existing deterministic helper summaries:
  `docs/specs/vba-effect-analysis.md` and
  `docs/adr/ADR-0023-procedure-effect-summaries.md`.
- Shared rule metadata and generated catalog contract:
  `docs/adr/ADR-0024-shared-static-analysis-rule-registry.md`.

## Related

- `docs/specs/excel-loop-performance-analysis.md`
- `docs/specs/cli-contract.md`
- `docs/adr/ADR-0013-analyze-runtime-risk-ownership.md`
- `docs/adr/ADR-0021-procedure-analysis-ir.md`
- `docs/adr/ADR-0022-conservative-vba-control-flow-graph.md`
- `docs/adr/ADR-0023-procedure-effect-summaries.md`
- `docs/adr/ADR-0024-shared-static-analysis-rule-registry.md`
- xlflow issue #452
- xlflow issue #453
- xlflow issue #454
- xlflow issue #455
