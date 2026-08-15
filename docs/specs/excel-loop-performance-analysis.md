# Excel Loop Performance Analysis

This specification defines the `VBA225`, `VBA238`, `VBA241`, `VBA242`, and
`VBA243` analyzer rules for Excel loop-performance issues `#452`, `#453`,
`#454`, `#455`, and `#456`.
`VBA225`
identifies conservative, repeated Excel object-model work inside VBA loops and
recommends bulk range or array operations. `VBA238` identifies loop-invariant
Excel object expressions that are repeatedly resolved and recommends hoisting
them into cached locals. `VBA241` identifies repeated `ReDim Preserve` array
copies inside loops and recommends preallocation or geometric growth. `VBA242`
identifies costly operations over entire rows, columns, worksheets, or
unbounded `UsedRange` expressions and recommends deriving bounded limits.
`VBA243` identifies bulk or repeated `Range.Value` transfers where `Value2`
may avoid unnecessary Currency or Date coercion while preserving semantics.
ADR-0025
records the design rationale and boundaries.

## Scope and public contract

<!-- xlflow-rule-contract: {"id":"VBA225","family":"analyze","category":"performance","default_severity":"warning","scope":"interprocedural","realtime":true,"configuration_key":"detect_excel_cell_access_in_loops","inline_suppressible":true,"preflight_blocking":false} -->

`VBA225` is a default-enabled `analyze` warning in batch analysis and the
shared real-time editor path. It is interprocedural when a uniquely resolved
project-local helper contributes the Excel access, remains non-blocking for
source preflight, and supports normal inline suppression. Its registry metadata
is:

- family: `analyze`;
- category: `performance`;
- default severity: `warning`;
- scope: `interprocedural`;
- precision: `medium`;
- configuration: `detect_excel_cell_access_in_loops`;
- inline suppression: supported;
- source-preflight blocking: not supported; and
- real-time editor diagnostics: supported.

The rule adds no fields to the `analysis` finding object. It uses the existing
`code`, `severity`, `file`, `module`, `procedure`, `line`, `message`, `reason`,
`suggestion`, and `nearby_code` fields. The message and reason identify the
access kind and explain that each iteration can add an Excel COM round trip.
Nested-loop findings retain `warning` severity and remain advisory.

## Loop boundaries and reachability

The analyzer consumes the normalized procedure IR and conservative CFG. It
recognizes `For`, `For Each`, `While`, and `Do` loops, including pre-test and
post-test forms, and considers only reachable loop-body statements. The loop
depth at the access site is retained for remediation text:

- every depth produces a `warning`; and
- depth two or greater adds context explaining that nesting multiplies the
  possible number of COM calls, without escalating severity.

One finding is emitted for each loop with confirmed hot-loop evidence. It is
located at the loop header, with deterministic source ordering. Additional
access kinds in the same loop are summarized in the reason rather than
producing duplicate findings for the same loop.

An indexed loop is eligible when its body contains a recognized Excel access,
even if the bound comes from a worksheet or another runtime value. A loop with
a statically provable upper bound of at most three iterations is exempt when
the bound is formed from integer literals or simple constant integer
arithmetic. A literal `For Each` range with at most three cells is also
trivial. An unknown bound, a data-dependent bound, or an unproven `For Each`
collection size is not treated as trivial.

## Excel access classification

The rule resolves the receiver and member through the existing VBA analysis
context where possible. It does not infer Excel identity from arbitrary text,
and ambiguous, unresolved, late-bound, or dynamically dispatched expressions
do not establish certain evidence.

| Access kind                 | Representative evidence                                                                                                  |
| --------------------------- | ------------------------------------------------------------------------------------------------------------------------ |
| Cell or range read          | `cell.Value`, `rng.Value2`, or `rng.Formula` consumed in an expression or assigned to a VBA value.                       |
| Cell or range write         | `cell.Value = ...`, `rng.Value2 = ...`, or `rng.Formula = ...` where the target is selected per iteration.               |
| Formatting                  | Per-cell `.Interior`, `.Font`, `.NumberFormat`, `.Borders`, alignment, or clear-format operations.                       |
| Address or worksheet lookup | Repeated `Cells`, `Range`, `Offset`, `Worksheets`, or `Sheets` lookup used to reach an iteration-specific cell or range. |
| Worksheet function call     | A resolved `WorksheetFunction` or Excel `Application` function called for the current loop item or cell.                 |

Unqualified Excel root names are recognized only when the procedure IR does not
contain a project-local binding with the same name. A local, parameter, or
module-level identifier such as `Dim cells As Variant` is therefore not
reclassified as Excel's implicit `Cells` function; its helper summary must not
create a `VBA225` finding either. A binding explicitly typed as `Range` or
`Worksheet` remains eligible, and the same binding rule is used by batch and
real-time analysis.

`For Each cell In range.Cells` is treated as a cell loop. Reads, writes,
formula assignments, formatting, and recognized worksheet-function calls on
`cell` are therefore eligible even when the loop variable is not used as an
explicit `Cells(row, column)` index. A helper call inside a loop contributes
evidence only when it resolves to one project-local procedure whose direct or
propagated summary contains one of the recognized access kinds. Ambiguous,
external, and dynamic helper calls remain uncertainty and do not create a
`VBA225` finding by themselves. Batch analysis uses project-wide summaries;
the real-time path uses summaries available in the current document snapshot.

## Bulk-operation exemptions and remediation

The rule is about repeated cell-by-cell or per-item object-model boundaries,
not every range operation. It does not report a range-level transfer or
formatting operation when the target and shape are independent of the current
loop item, for example:

```vb
values = ws.Range("A1:A1000").Value2
ws.Range("B1:B1000").Value2 = values
ws.Range("A1:A1000").NumberFormat = "0.00"
```

An expression rooted at the current `cell`, or an address built from the
current loop variable, is not a bulk operation even if it uses `.Value2` or
`.Formula`. The suggestion is access-kind specific:

- reads should load one rectangular range into a VBA array, process the array,
  and read Excel once;
- writes and formulas should populate an array and assign it to one range;
- formatting should target the complete range or a small set of contiguous
  ranges; and
- repeated worksheet or range lookup should be hoisted and cached when the
  root is invariant, while the per-cell work is moved into the array pass.

The diagnostic explains that COM calls made once per iteration can dominate
runtime even when the VBA body is small. It does not claim a fixed slowdown or
recommend suppressing the rule solely because a loop is currently fast.

## Configuration and compatibility

Projects can disable the rule with the stable ID:

```toml
[analyze]
disabled_rules = ["VBA225"]
```

The legacy Boolean `detect_excel_cell_access_in_loops = false` remains accepted
for compatibility and emits the normal deprecation warning. An inline
`xlflow:disable-line VBA225` or `xlflow:disable-next-line VBA225` comment may
be used for an intentional, locally bounded exception.

`VBA225` remains separate from `VBA205`: the latter reports ambiguous workbook
or worksheet scope, while this rule reports repeated, resolved Excel work
inside a loop. It also remains separate from correctness and reliability rules
such as `VBA215` and `VBA218`.

The generated diagnostic catalog must be regenerated from the static-analysis
registry after the registry implementation is updated. The generated
`vitepress/reference/diagnostics.md` file is not a hand-edited source of truth.

## Loop-invariant Excel object resolution (`VBA238`)

<!-- xlflow-rule-contract: {"id":"VBA238","family":"analyze","category":"performance","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_loop_invariant_excel_object_resolution","inline_suppressible":true,"preflight_blocking":false} -->

`VBA238` is a default-enabled, warning-level performance diagnostic for
loop-invariant Excel object-model resolution. It reports a repeated member
chain whose receiver and constant arguments do not depend on the active loop
iterator, such as `ThisWorkbook.Worksheets("Data")`,
`Workbooks("Book.xlsx").Worksheets("Data")`, `ListObjects("Sales")`, named
ranges, pivot tables, or charts. The rule is available in batch analysis and
the shared real-time editor path, remains non-blocking, and supports normal
inline suppression.

The analyzer normalizes equivalent member chains across whitespace, line
continuations, and `With` blocks before comparing them. It ignores simple local
variable access and expressions that reference the loop variable (for example,
`Cells(i, 1)`), because those values must be resolved per iteration. Nested
loops are supported: an expression is eligible only when it is invariant with
respect to the loop in which it is reported. Ambiguous or unresolved chains do
not establish a finding.

The suggested remediation extracts the invariant chain into a local object
outside the loop and reuses it:

```vb
Dim dataSheet As Worksheet
Set dataSheet = ThisWorkbook.Worksheets("Data")
For i = 1 To lastRow
    dataSheet.Cells(i, 1).Value2 = i
Next
```

Projects can disable the rule with:

```toml
[analyze]
disabled_rules = ["VBA238"]
```

An intentional local exception can use `xlflow:disable-line VBA238` or
`xlflow:disable-next-line VBA238`. `VBA238` is separate from `VBA225`:
`VBA225` reports the repeated per-cell or per-item Excel work itself, whereas
`VBA238` reports the avoidable repeated lookup of an invariant object chain.

## Repeated `ReDim Preserve` in loops (`VBA241`)

<!-- xlflow-rule-contract: {"id":"VBA241","family":"analyze","category":"performance","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_redim_preserve_in_loops","inline_suppressible":true,"preflight_blocking":false} -->

`VBA241` is a default-enabled, procedure-local performance diagnostic in batch
analysis and the shared real-time editor path. It reports reachable
`ReDim Preserve` statements in `For`, `For Each`, `While`/`Wend`, and all
supported `Do`/`Loop` forms, including pre-test and post-test `While`/`Until`
conditions. Synthetic condition nodes are not treated as loop bodies. Fixed
arrays and scalar targets remain owned by the existing allocation and
correctness rules rather than producing a `VBA241` finding.

The rule reuses the existing `ReDim Preserve` parser, target splitting, and
dimension-expression analysis. It compares variable accesses within each
dimension-bound expression with the control variables of all containing loops,
so direct references and helper expressions such as `Grow(i)`, `i + Offset()`,
and parenthesized forms are classified consistently. Helper procedure bodies
are not analyzed transitively; only accesses visible in the dimension
expression establish loop-variable dependence. All dimensions are considered
for multidimensional arrays, while non-final-dimension correctness remains
owned by `VBA208`.

One finding is emitted at the `ReDim Preserve` statement even when a statement
contains multiple targets. A single loop with dimensions independent of every
containing loop variable is classified as repeated constant-size reallocation
and uses `information` severity. A dimension that depends on the current or an
outer loop variable is classified as growth based on a loop variable and uses
`warning` severity. Any finding at nesting depth two or greater also uses
`warning`, with the depth included in the message and reason. Severity never
exceeds `warning`.

Suggestions distinguish the remediation: preallocate the final array size when
the bound is known or can be computed before the loop; otherwise maintain a
capacity variable and grow geometrically before copying or assigning new
items. The rule is advisory, non-blocking for source preflight, and supports
normal inline suppression. Disable it with:

```toml
[analyze]
disabled_rules = ["VBA241"]
```

The compatibility key `detect_redim_preserve_in_loops = false` remains accepted
with the normal deprecation warning. The generated diagnostic catalog is
regenerated from the static-analysis registry and is not edited by hand.

## Expensive full-range operations (`VBA242`)

<!-- xlflow-rule-contract: {"id":"VBA242","family":"analyze","category":"performance","default_severity":"information","scope":"procedure-local","realtime":true,"configuration_key":"detect_expensive_full_range_operations","inline_suppressible":true,"preflight_blocking":false} -->

`VBA242` is an opt-in, procedure-local performance diagnostic in batch
analysis and the shared real-time editor path. It recognizes Excel range
expressions whose shape covers an entire row, column, or worksheet, including
`Rows`, `Columns`, `EntireRow`, `EntireColumn`, a bare `Cells`, and direct or
unbounded `UsedRange` receivers. It reports only when one of the expressions
feeds a costly operation such as a `Value`/`Formula` assignment, calculation,
formatting, `Find`, `Replace`, or sorting. A range expression by itself is not
reported.

The rule supports `information` outside loops and escalates to `warning` when
the operation is reachable from a `For`, `For Each`, `While`, or `Do` loop.
The finding remains advisory, non-blocking, and inline suppressible. Explicit
bounded A1 ranges, `Range(Cells(...), Cells(...))`, a bounded `Resize`, and an
`Intersect` with a bounded range are accepted. The v1 implementation does not
infer effective bounds from a guard in another statement, a named range, or a
range alias whose shape is unknown.

The rule is disabled by default because one-time whole-sheet formatting is a
legitimate Excel operation. Enable it for an Excel project with:

```toml
[analyze]
detect_expensive_full_range_operations = true
```

Prefer `[analyze].disabled_rules = ["VBA242"]` for explicit policy and use
`xlflow:disable-line VBA242` or `xlflow:disable-next-line VBA242` for a local,
intentional whole-range operation. The compatibility Boolean emits the normal
deprecation warning. `VBA242` owns only oversized operation targets; `VBA201`
continues to own unsafe `Find` result use, `VBA217` owns last-row guidance, and
`VBA225` owns repeated per-cell Excel work.

## Value2 performance opportunities (`VBA243`)

<!-- xlflow-rule-contract: {"id":"VBA243","family":"analyze","category":"performance","default_severity":"information","scope":"procedure-local","realtime":true,"configuration_key":"detect_value2_performance_opportunities","inline_suppressible":true,"preflight_blocking":false} -->

`VBA243` is an opt-in, procedure-local performance diagnostic in batch
analysis and the shared real-time editor path. It identifies a bulk or
repeated `Range.Value` transfer with a strong performance signal, such as a
large or dynamic range, repeated access inside a loop, or immediate transfer to
a `Variant` array. Numeric/text processing can strengthen the explanation
after one of those signals exists, but is not a standalone trigger. It
suggests using `Range.Value2` when the raw-value semantics are acceptable; it
does not globally prefer `Value2`
or report an intentional `Value` use when date or currency handling is clear.
Literal ranges are considered large at 100 or more cells. Dynamic evidence
requires a proven runtime construction such as concatenated bounds, a
`Range(Cells(...), Cells(...))` pair, `Resize`, `CurrentRegion`, `UsedRange`,
or a traceable alias; an otherwise unknown `Range` argument or variable does
not qualify by itself. Typed `Date`/`Currency` values, date literals, conversion
or formatting calls, subtype inspection, and mixed or unknown Variant uses are
conservatively excluded.

The rule supports `information` outside loops and may use `warning` when a
transfer is repeated from a reachable loop. It remains advisory,
non-blocking, and inline suppressible. Disable it with:

```toml
[analyze]
disabled_rules = ["VBA243"]
```

The compatibility key `detect_value2_performance_opportunities = true`
remains accepted with the normal deprecation warning. Use
`xlflow:disable-line VBA243` or `xlflow:disable-next-line VBA243` for a local
exception. `VBA243` owns only the `Value` versus `Value2` performance
opportunity; `VBA226` continues to own unsafe value-array shape assumptions.

## Related

- `docs/adr/ADR-0025-excel-loop-performance-analysis.md`
- `docs/adr/ADR-0013-analyze-runtime-risk-ownership.md`
- `docs/adr/ADR-0021-procedure-analysis-ir.md`
- `docs/adr/ADR-0022-conservative-vba-control-flow-graph.md`
- `docs/adr/ADR-0023-procedure-effect-summaries.md`
- `docs/adr/ADR-0024-shared-static-analysis-rule-registry.md`
- xlflow issue #452
- xlflow issue #453
- xlflow issue #454
- xlflow issue #455
- xlflow issue #456
