# VBA Dead-Store Diagnostics

<!-- xlflow-rule-contract: {"id":"VBA256","family":"analyze","category":"reliability","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_dead_stores","inline_suppressible":true,"preflight_blocking":false} -->

`VBA256` reports a procedure-local scalar assignment whose resulting value is
not read on any path before another assignment replaces it or the procedure
exits. The rule diagnoses the unused assigned value, not the whole statement:
the right-hand side can still have observable effects and must not be removed
without separate review.

## Eligible storage and assignments

The first implementation intentionally has a narrow, high-precision scope. A
candidate must be a simple identifier target resolved to a non-`Static` local
declaration with an explicit scalar type. The assignment must be an ordinary
implicit or explicit `Let` assignment whose target access is write-only.

The rule does not report assignments to parameters, procedure return slots,
module or project variables, unresolved identifiers, `Variant`, arrays,
objects, properties, default members, indexed storage, `Static` locals, loop
control variables, `Set` targets, or `ReDim` targets. Read/write operations are
not reduced to write-only candidates.

## Control-flow contract

The analyzer computes backward liveness over the existing conservative
procedure CFG. Branches, loops, labels, `GoTo`, error handlers, `Resume`, and
all synthetic exits retain their existing CFG meaning. If any reachable path
can read the assigned value before replacement or exit, no finding is emitted.

A normal successor observes state after a completed assignment. An
exceptional successor from the assignment statement observes the state before
that statement completed and therefore cannot by itself observe the new
value. Uncertain normal flow, recovered control flow, or a reachable unknown
destination prevents a dead-store proof for affected values.

The procedure CFG does not choose an active conditional-compilation branch.
If a local declaration or any access to that local appears inside
`#If`/`#ElseIf`/`#Else`, the rule excludes that local from dead-store analysis
rather than infer liveness across mutually exclusive build configurations.

A comparison expression nested in an assignment's value contributes only
reads. Its operands are never the assignment target, even when the target is
a member expression rather than a local identifier.

Arguments of a statement-level call that the parser could only recover as a
leaf `call_statement` — for example a parenthesis-free call whose first
argument is an implicit `With` member — still record identifier reads, so a
local passed to such a call is observed.

A single-line `If` owns every statement after `Then` on its logical line. If a
same-line sibling still follows it, the parser could not represent the line's
control flow and the procedure's CFG is untrustworthy, so the rule reports no
findings for that procedure.

Calls observe direct scalar argument values. A local passed through a form
that may be `ByRef`, or through an unresolved, ambiguous, external, or dynamic
call whose argument contract cannot be proven, remains live at that call. The
rule does not add interprocedural mutation summaries.

## Surfaces and configuration

`VBA256` is an opt-in warning-level, non-blocking, high-precision analyzer rule
available in batch and realtime/LSP analysis. Enable it with
`detect_dead_stores = true`, or add `VBA256` to `[analyze].disabled_rules` after
enabling it, or use an
inline `xlflow:disable-line VBA256` / `xlflow:disable-next-line VBA256`
suppression.

`VB020` continues to own declaration-level unused-local findings. `VBA256`
owns assignment-level value liveness even when both findings can point to
different remediation decisions.

## Precision and performance

The rule fails open when storage identity, assignment completion, or control
flow cannot be established. It performs at most one procedure-local fixed
point for a procedure containing an eligible scalar assignment. Procedures
without scalar assignment syntax do not enter the liveness walk. Diagnostic
ordering follows the analyzer's deterministic source-order sort.
