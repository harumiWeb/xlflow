# ADR-0024: Procedure-Local VBA Resource Lifetime Analysis

## Status

Accepted

## Context

The conservative CFG in ADR-0022 can establish whether cleanup is reached on
normal, exceptional, termination, and unknown exits. It does not itself model
which procedure acquires a resource, whether a local alias identifies that
resource, or whether an object is borrowed from a caller. Issue #437 needs a
high-confidence resource-leak diagnostic without treating every VBA object
reference as owned.

## Decision

Add a procedure-local analyzer rule, `VBA219`, for explicit Workbook and VBA
file-handle acquisitions. It uses the existing normalized procedure IR and CFG;
it does not add a new CFG or interprocedural effect layer.

Each explicit acquisition is tracked independently from its normal successor.
An acquisition must be released before every reachable CFG exit. Direct local
aliases can identify the same resource for Close. Parameters and aliases that
originate only from parameters are borrowed and never create an ownership
obligation. A local Workbook may transfer ownership only when its alias remains
in an object-returning Function's return slot through a successful normal
Function exit.

Release means a recognized `Workbook.Close`, `Close #handle`, or bare `Close`
statement is reached. The rule intentionally does not model a Close call's own
failure, arbitrary helper calls, ByRef ownership transfer, or dynamic dispatch.

## Consequences

- The analyzer emits a default-enabled, non-blocking, suppressible warning at
  the acquisition site and names an uncovered exit witness.
- Batch analysis and realtime diagnostics reuse the same procedure-local
  implementation and CFG semantics.
- The first release vocabulary is deliberately narrow: FSO, ADODB, Office
  automation, temporary files, and cross-procedure ownership remain future
  work.

## Evidence

- CFG exit and cleanup contract: `docs/specs/vba-control-flow-graph.md` and
  `internal/vba/cfg`.
- Analyzer ownership and public diagnostic policy:
  `docs/adr/ADR-0013-analyze-runtime-risk-ownership.md` and
  `internal/analyze`.
- Procedure IR ownership and local declaration contract:
  `docs/adr/ADR-0021-procedure-analysis-ir.md` and
  `internal/vba/procedureir`.

## Related

- `docs/adr/ADR-0013-analyze-runtime-risk-ownership.md`
- `docs/adr/ADR-0021-procedure-analysis-ir.md`
- `docs/adr/ADR-0022-conservative-vba-control-flow-graph.md`
- `docs/specs/vba-resource-lifecycle-analysis.md`
- xlflow issue #437
