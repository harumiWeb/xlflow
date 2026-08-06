# Excel Loop Performance Analysis

This specification defines the `VBA225` analyzer rule for issue #452. The
rule identifies conservative, repeated Excel object-model work inside VBA
loops and recommends bulk range or array operations. ADR-0025 records the
design rationale and boundaries.

## Scope and public contract

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
Nested-loop findings use `error` severity, but remain advisory and do not
change the source-preflight policy.

## Loop boundaries and reachability

The analyzer consumes the normalized procedure IR and conservative CFG. It
recognizes `For`, `For Each`, `While`, and `Do` loops, including pre-test and
post-test forms, and considers only reachable loop-body statements. The loop
depth at the access site is retained for severity and remediation text:

- depth one produces a `warning`; and
- depth two or greater produces an `error` because the nested loop multiplies
  the possible number of COM calls.

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

## Related

- `docs/adr/ADR-0025-excel-loop-performance-analysis.md`
- `docs/adr/ADR-0013-analyze-runtime-risk-ownership.md`
- `docs/adr/ADR-0021-procedure-analysis-ir.md`
- `docs/adr/ADR-0022-conservative-vba-control-flow-graph.md`
- `docs/adr/ADR-0023-procedure-effect-summaries.md`
- `docs/adr/ADR-0024-shared-static-analysis-rule-registry.md`
- xlflow issue #452
