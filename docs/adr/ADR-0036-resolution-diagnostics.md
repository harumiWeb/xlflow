# ADR-0036: Canonical Resolution Outcomes for Compile-Equivalent Diagnostics

## Status

Accepted

## Context

Call hierarchy, impact, effect propagation, language-server features, and
batch analysis historically projected project symbols through more than one
resolver. That made it possible for visibility, receiver, and incomplete
workspace decisions to diverge. Issue #590 also needs to distinguish a
provably invalid call from an external, late-bound, dynamic, or incomplete
reference without turning uncertainty into a compile error.

## Decision

Extend the `procedureir` resolver from a procedure-only result to the shared
outcome model `matched`, `non_callable`, `ambiguous`, `unresolved`, `external`,
`builtin_like`, `dynamic`, and `incomplete`. The resolver owns candidate
ordering, module/receiver visibility, lexical shadowing, enum-member lookup,
and same-object-module event lookup. Enum members and `RaiseEvent` references
are retained in the IR with source ranges, conditional and recovery facts.

`calls.Resolver` is a compatibility projection over that resolver. Legacy
inspect-calls status values and JSON remain stable; richer negative states are
projected to the existing unresolved/member representations for graph clients.
Call hierarchy, impact, effects, and LSP workspace snapshots use the same
resolver snapshot.

The diagnostic policy emits the compile-equivalent `VB052`, `VB053`, and
`VB054` only when the project model is complete and the outcome proves an
invalid project-local call, an ambiguous bare Enum member, or an undeclared
same-module event. Parser recovery, unresolved conditional branches,
filtered/partial projects, unavailable TypeLib data, external members, and
dynamic APIs fail open. LSP Fast diagnostics and incomplete Full snapshots
exclude these project-negative diagnostics.

## Consequences

- Batch and LSP diagnostics share the same candidate and completeness
  semantics, while existing call-graph uncertainty behavior is preserved.
- The resolver and IR carry additional metadata, so consumers must clone and
  rebase enum/event facts along with procedure calls.
- A stale or recovered workspace may temporarily suppress a negative
  diagnostic until a clean complete snapshot is available.
- TypeLib constant lookup can retain all candidates through the diagnostic
  API while preserving the historical single-winner `ResolveConstant` API.

## Alternatives considered

1. Add a separate lint-only name matcher. Rejected because it would diverge
   from call hierarchy, effects, and LSP visibility semantics.
2. Report every unresolved bare name. Rejected because the target may be an
   external library member, late-bound value, or incomplete project snapshot.
3. Replace the legacy calls JSON status model. Rejected because inspect and
   downstream graph consumers require compatibility.

## Evidence

- `internal/vba/procedureir/resolver.go` and focused resolver/IR tests.
- `internal/vba/calls/calls.go` compatibility adapter tests.
- LSP workspace completeness regression tests.
- Registry contracts for `VB052` through `VB054` and the shared diagnostics
  policy.

## Related

- Issue #590
- ADR-0021, ADR-0023, ADR-0024
