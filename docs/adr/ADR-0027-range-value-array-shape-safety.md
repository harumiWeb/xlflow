# ADR-0027: Conservative Range.Value Array-Shape Safety Analysis

## Status

Accepted

## Context

Excel returns a scalar Variant for a one-cell `Range.Value` or `Range.Value2`
read, but returns a two-dimensional Variant array for a multi-cell read. The
shape is independent of whether the range is vertical, horizontal, or
rectangular. VBA code commonly assumes that a column or row produces a
one-dimensional array, calls `UBound` without a dimension, or writes the result
to a destination range whose shape is not compatible.

A text-only check cannot distinguish a definite single-cell range from a
multi-cell range, cannot follow a Range alias, and cannot preserve safety after
branch joins or reassignment. The analyzer already has procedure IR, CFG, and
the shared rule registry needed to add a procedure-local runtime-safety rule
without changing the public finding envelope.

## Decision

Add `VBA226` as a default-enabled, warning-level, non-blocking, inline-
suppressible, real-time `analyze` rule. The rule is configured by
`detect_range_value_array_shape` and uses the existing `Finding` JSON shape.

Track only procedure-local values obtained from `Range.Value` or `Range.Value2`.
Represent their possible shape as a definite scalar, a definite two-dimensional
array with known row and column counts, or an uncertain value. Infer literal
`Range` and `Cells` shapes and simple static Range aliases; dynamic expressions,
reassignments, and incompatible branch joins become uncertain.

Treat one-index access and dimensionless `LBound`/`UBound` as unsafe for tracked
Range values. Accept two-dimensional access when the value is proven to be an
array, the access has row-first/column-second dimensions, and statically known
index bounds fit the source shape. Recognize `IsArray` true branches and
early-exit false guards, but do not let `IsArray` make a one-dimensional access
safe. Diagnose a block assignment only when a definite two-dimensional source
array is written to a proven incompatible destination shape; unchanged
pass-throughs and equal-shape transfers remain quiet.

The rule consumes the normalized procedure IR and conservative normal CFG
paths. It does not perform interprocedural shape propagation, open Excel, or
block source preflight.

## Consequences

- Common vertical, horizontal, rectangular, and dynamic-range mistakes are
  reported before execution with actionable row/column guidance.
- Conservative joins and dynamic expressions may report an uncertain scalar or
  array use, but explicit `IsArray` guards provide a local proof for safe
  two-dimensional access.
- The rule must keep its registry, configuration adapter, CLI documentation,
  generated catalog, and batch/LSP implementations synchronized.
- It intentionally does not prove arbitrary runtime index arithmetic or helper
  return shapes; those cases remain outside the procedure-local contract.

## Alternatives Considered

1. **Search for `Value2` and parentheses in source text** - Rejected because
   it cannot infer range shape, path-sensitive guards, or destination
   compatibility and would produce false positives for comments and strings.
2. **Treat every Range.Value result as an array** - Rejected because a
   single-cell read is a scalar and must remain valid scalar VBA code.
3. **Treat every Range.Value result as a scalar** - Rejected because it would
   miss safe and unsafe two-dimensional block processing.
4. **Block push/run preflight for uncertain shapes** - Rejected because the
   rule is advisory runtime-safety feedback and cannot prove workbook data or
   dynamic range sizes statically.

## Evidence

- Issue #443 requirements and acceptance criteria.
- Analyzer ownership: `docs/adr/ADR-0013-analyze-runtime-risk-ownership.md`.
- Procedure IR and CFG boundaries: `docs/adr/ADR-0021-procedure-analysis-ir.md`
  and `docs/adr/ADR-0022-conservative-vba-control-flow-graph.md`.
- Shared rule metadata: `docs/adr/ADR-0024-shared-static-analysis-rule-registry.md`.
- Implementation and tests: `internal/analyze/range_value_shape.go` and
  `internal/analyze/range_value_shape_test.go`.

## Related

- `docs/specs/range-value-array-shape-safety.md`
- `docs/specs/cli-contract.md`
- `vitepress/reference/diagnostics.md`
