# Range.Value Array-Shape Safety

This specification defines the `VBA226` analyzer rule for issue #443. The
rule detects scalar and one-dimensional assumptions made about `Range.Value` or
`Range.Value2` results.

## Public contract

`VBA226` is a default-enabled `analyze` warning in the runtime-safety category.
It is procedure-local, medium precision, non-blocking, inline-suppressible,
and available in batch and real-time analysis. Its configuration key is
`detect_range_value_array_shape`; projects may also disable it with:

```toml
[analyze]
disabled_rules = ["VBA226"]
```

The rule uses the existing finding fields and adds no JSON fields, CLI flags,
or LSP capabilities.

## Shape model

- A definite one-cell range is a scalar Variant.
- A definite multi-cell range is a two-dimensional, one-based Variant array.
- Vertical, horizontal, and rectangular ranges all use row-first,
  column-second indexing.
- Dynamic ranges, unknown Range aliases, reassignments, and branch joins with
  different shapes are uncertain.

Literal A1 ranges, two-literal `Range(start, end)` calls, literal `Cells(row,
column)` calls, and simple `Set` Range aliases may be resolved. Arbitrary
runtime expressions are not resolved into a guessed size.

## Diagnostics and safe forms

The rule reports:

- `values(index)` for a tracked Range value;
- `LBound(values)` or `UBound(values)` without a dimension;
- two-dimensional access on a definite scalar or an unguarded uncertain value;
- statically proven row/column bounds or dimension order violations; and
- assignment of a definite two-dimensional array to a proven incompatible
  destination range.

The following are accepted when their bounds are not statically disproven:

```vb
values = ws.Range("A1:B10").Value2
For rowIndex = LBound(values, 1) To UBound(values, 1)
    For columnIndex = LBound(values, 2) To UBound(values, 2)
        Debug.Print values(rowIndex, columnIndex)
    Next columnIndex
Next rowIndex
```

`IsArray(values)` and an early-exit scalar guard establish that a dynamic
Range.Value result is in its two-dimensional branch. They do not make
`values(index)` safe because a Range array is not one-dimensional.

Direct block transfers are quiet when the source and destination are proven to
have the same shape. Unknown shapes are not reported unless the code makes a
definite incompatible-shape claim; the rule does not warn about unchanged
pass-through values.

## Boundaries

The rule is procedure-local and does not infer shapes through helper return
values or interprocedural calls. It ignores comments and string literals,
does not open Excel, and does not block `push` or `run` preflight.

## Related

- `docs/adr/ADR-0027-range-value-array-shape-safety.md`
- `docs/specs/cli-contract.md`
- Issue #443
