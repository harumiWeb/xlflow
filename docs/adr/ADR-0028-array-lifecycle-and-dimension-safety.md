# ADR-0028: Common VBA Array Lifecycle and Dimension Safety Model

## Status

Accepted

## Context

The analyzer previously implemented `VBA208` (`ReDim Preserve`) and `VBA209`
(array comparison) as independent text checks. `VBA226` introduced a separate
procedure-local state for `Range.Value` shapes. That separation cannot answer
whether a dynamic array is allocated after a branch or `Erase`, cannot retain
known dimensions through `ReDim`, and makes it easy for access diagnostics to
disagree with existing array rules.

VBA allocation is path-sensitive: fixed arrays are available at procedure
entry, dynamic arrays are not, `Erase` has different effects for fixed and
dynamic arrays, and a Variant can change between scalar and array values.
Project-local functions can also return arrays, but only a unique and
consistent return contract is safe to use. The public finding envelope and
existing diagnostic IDs must remain compatible.

## Decision

Add a shared CFG-based array state transfer in `internal/analyze`. Each tracked
value carries `allocated`, `unallocated`, or `unknown` state together with
array kind, element/object/Variant information, known dimension count, known
bounds, and origin metadata. Merge operations retain allocated state and shape
only when all incoming paths agree; exceptional and uncertain edges propagate
the pre-statement state.

Add default-enabled `VBA227` for lifecycle, access, dimension, bound, and
unknown-Variant safety. Keep `VBA208`, `VBA209`, `VBA101` / `VBA102`, and
`VBA226` as public contracts, but issue their applicable findings from the
shared state or explicitly preserve their existing ownership. In particular,
`VBA226` continues to own `Range.Value` / `Value2` shape diagnostics and the
common model marks those values as Range-origin to prevent duplicates.

Summarize only unique, consistent project-local Function and Property Get
array returns in batch analysis; real-time analysis uses the same-document
subset. Mixed, recursive, ambiguous, and external returns are unknown.
`IsArray` can establish an array kind in a future guard refinement but does not
by itself establish allocation or dimensions. The rule is warning-level,
inline-suppressible, non-blocking, and keeps the existing JSON and CLI
contracts.

## Consequences

- Array access, bounds, `ReDim`, `Erase`, and object-array assignment checks use
  one conservative lifecycle model.
- Branch joins and exceptional paths may produce warnings when allocation is
  not proven on every path; this is intentional runtime-safety behavior.
- Known project-local array returns can avoid false warnings without trusting
  external or ambiguous calls.
- `VBA226` remains independently testable and does not produce duplicate
  lifecycle findings for Range-origin values.
- The rule registry, configuration adapters, realtime implementation list,
  CLI/specification docs, generated diagnostic references, and batch/realtime
  tests must be updated together.

## Alternatives Considered

1. **Add one text matcher per target pattern** - Rejected because it cannot
   model allocation across CFG joins, `Erase`, Variant reassignment, or known
   dimensions and would duplicate existing diagnostics.
2. **Treat every dynamic array access as unsafe** - Rejected because it would
   warn after an unconditional `ReDim` and miss useful shape/bound evidence.
3. **Treat unknown Variant and external results as allocated arrays** -
   Rejected because it converts uncertainty into an unsafe runtime guarantee.
4. **Fold all Range.Value shape logic into VBA227** - Rejected because
   `VBA226` already owns that public contract and its shape-specific tests and
   diagnostics should remain stable.

## Evidence

- Issue #444 requirements and acceptance criteria.
- Analyzer ownership: `docs/adr/ADR-0013-analyze-runtime-risk-ownership.md`.
- Procedure IR and CFG boundaries: `docs/adr/ADR-0021-procedure-analysis-ir.md`
  and `docs/adr/ADR-0022-conservative-vba-control-flow-graph.md`.
- Shared metadata contract: `docs/adr/ADR-0024-shared-static-analysis-rule-registry.md`.
- Range shape predecessor: `docs/adr/ADR-0027-range-value-array-shape-safety.md`.
- Implementation and regression coverage: `internal/analyze/array_safety.go`
  and `internal/analyze/analyzer_test.go`.

## Supersedes

None.

## Superseded by

None.

## Related

- `docs/specs/array-lifecycle-safety.md`
- `docs/specs/range-value-array-shape-safety.md`
- `docs/specs/cli-contract.md`
- `docs/adr/ADR-0027-range-value-array-shape-safety.md`
