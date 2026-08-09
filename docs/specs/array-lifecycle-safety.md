# Array Lifecycle and Dimension Safety

This specification defines the `VBA227` analyzer rule for issue #444. It
extends the existing `ReDim Preserve`, array comparison, object-array `Set`,
and `Range.Value` checks through one conservative array-state model.

## Public contract

`VBA227` is a default-enabled `analyze` warning in the runtime-safety category.
It is interprocedural for unique project-local array-return summaries,
medium-precision, non-blocking, inline-suppressible, and available in batch and
real-time analysis. Its configuration key is
`detect_array_lifecycle_safety`; projects may also disable it with:

```toml
[analyze]
disabled_rules = ["VBA227"]
```

Disabling `VBA227` suppresses only findings issued under `VBA227`. The shared
array state still supplies independently controlled `VBA208` and `VBA209`
findings, and it can still supply the always-enabled object-array `VBA101` /
`VBA102` findings. Disable those rules separately when that compatibility
behavior is intended; `VBA227` is not the switch for `ReDim Preserve` or array
comparison safety.

The rule uses the existing `Finding` fields and adds no JSON fields, CLI flags,
or LSP capabilities.

## State model

The analyzer carries a case-insensitive variable state through the procedure
CFG. Allocation is a three-point lattice:

- `allocated`: the value is proven to be an array allocation on every incoming
  path;
- `unallocated`: a dynamic array is proven not to have an allocation; and
- `unknown`: a branch, exceptional edge, parameter, external result, or
  Variant assignment prevents a safe proof.

The state also records whether the value is a fixed or dynamic array, whether
it is an object array or Variant, the known dimension count, statically known
lower and upper bounds, and whether its origin is `Range.Value` / `Value2`.
Fixed-size arrays begin allocated. Dynamic arrays begin unallocated. `ReDim`
records its statically known dimensions and allocates a dynamic array. `Erase`
deallocates a dynamic array but preserves fixed-array allocation; Erase on a
fixed array is an element reset. Array literals, known array-returning
functions, and known array assignments establish allocation. Unknown Variant
and external values do not.

At a join, allocation and dimensions are retained only when all incoming paths
agree. Exceptional and uncertain CFG edges use the pre-statement state. This
means an indexed access after an allocation on only one branch remains a
warning.

## Diagnostics and ownership

`VBA227` reports:

- `LBound` / `UBound` before allocation, invalid dimension numbers, and
  dimension or bound contradictions;
- indexed access before `ReDim`, after `Erase`, or after an unproven Variant
  assignment;
- `ReDim` on fixed-size arrays or non-array values;
- known subscript-count mismatches and inconsistent loop lower-bound
  assumptions; and
- unknown Variant array operations.

`VBA208` remains the owner of `ReDim Preserve` safety. It warns when a
non-final dimension changes and remains conservative when the prior shape is
unknown. A one-dimensional target has no non-final dimension and is not
reported merely because its previous bounds are unavailable. For
multi-dimensional arrays, CFG joins retain each bound that agrees across the
incoming paths, and equivalent non-literal bound expressions such as repeated
`rowCount` references establish the same non-final shape. A `ReDim` of an
indexed member such as `items(i).values(...)` is not attributed to the receiver
array `items`; nested member-array lifecycle is outside the direct-variable
state model. `VBA209` remains the owner of scalar object comparisons and
receives array comparison findings from the common model. For arrays, VBA209
applies only when an array identifier is a direct operand of a comparison
expression; an array assignment, function or bound-call argument such as
`LBound(values)`, an indexed element, and a member access are not scalar array
comparisons. A parenthesized array identifier remains the same direct operand.
The common model also supplies object-array element assignments without `Set`
to the existing `VBA101` / `VBA102` finding constructors. `VBA226` remains the
owner of `Range.Value` / `Value2` shape diagnostics; Range-origin values are
excluded from duplicate `VBA227` access and bound findings.

## Function and property summaries

Batch analysis may use a unique project-local `Function` or `Property Get`
summary when every observed normal return assignment returns an allocated
array with a consistent shape. Real-time analysis restricts this summary to
the active document; it does not resolve array-return summaries from another
module. Mixed return kinds, missing assignments, recursion, ambiguous names,
and external calls remain unknown.

## Boundaries

The rule does not open Excel or claim that `IsArray` proves allocation or a
dimension count. It accepts literal numeric bounds only when they are
statically visible; dynamic expressions remain unknown. It does not change
the `Finding` JSON envelope or preflight blocking behavior.

## Related

- `docs/adr/ADR-0028-array-lifecycle-and-dimension-safety.md`
- `docs/adr/ADR-0027-range-value-array-shape-safety.md`
- `docs/specs/range-value-array-shape-safety.md`
- `docs/specs/cli-contract.md`
- Issue #444
