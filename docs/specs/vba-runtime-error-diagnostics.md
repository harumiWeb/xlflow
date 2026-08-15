# Deterministic VBA Runtime-Error Diagnostics

This specification defines `VBA249` for issue #596. It reports runtime
failures only when the shared constant, type, CFG, and dataflow facts prove
that the failure occurs at the expression on every relevant reachable path.
It does not interpret VBA, call Excel, or convert an unknown value into a
diagnostic. The architectural rationale is recorded in ADR-0042.

## Public contract

<!-- xlflow-rule-contract: {"id":"VBA249","family":"analyze","category":"runtime-safety","default_severity":"error","scope":"procedure-local","realtime":true,"configuration_key":"detect_deterministic_runtime_errors","inline_suppressible":true,"preflight_blocking":false} -->

`VBA249` is a default-enabled `analyze` error in the runtime-safety category.
It is procedure-local, high-precision, available in batch and real-time
analysis, non-blocking for source preflight, and inline-suppressible. Its
configuration key is `detect_deterministic_runtime_errors`; projects may also
disable it with:

```toml
[analyze]
disabled_rules = ["VBA249"]
```

The rule uses the existing finding envelope and adds an additive
`runtime_error` context. The stable field is `runtime_error.kind`; the initial
values are listed below. No CLI flag, JSON envelope version, or LSP capability
is added. An `error` severity here means that the runtime failure is proven; it
does not mean that `push`, `run`, or another workbook command is blocked before
Excel opens. `compile-equivalent` diagnostics remain the separate
preflight-blocking class described by the CLI contract.

## Deterministic failure kinds

| Kind                            | Report when                                                                                                        | Remain silent when                                                                                                          |
| ------------------------------- | ------------------------------------------------------------------------------------------------------------------ | --------------------------------------------------------------------------------------------------------------------------- |
| `division_by_zero`              | `/`, `\`, or `Mod` has a divisor proven equal to numeric zero.                                                     | The divisor is unknown, differs across paths, or is only possibly zero.                                                     |
| `numeric_type_mismatch`         | A known nonnumeric string or other known incompatible scalar is consumed by a supported numeric operator.          | The value is an unknown `Variant`, a locale-dependent coercion, or a user-defined/external value.                           |
| `conversion_type_mismatch`      | A supported explicit numeric conversion is applied to a known value whose type/value domain cannot be converted.   | The conversion or input depends on an unresolved call, unknown `Variant`, locale, or host-specific behavior.                |
| `conversion_overflow`           | A supported explicit numeric conversion is applied to a known value outside the target type's representable range. | The target range, source value, or platform width is not statically fixed.                                                  |
| `array_subscript_out_of_bounds` | Every relevant path gives a known lower/upper bound and the subscript is proven outside that bound.                | Any dimension, bound, subscript, `Option Base`, or array shape is unknown.                                                  |
| `array_unallocated`             | A dynamic array is proven unallocated on every relevant path reaching an index or bound access.                    | Allocation is possible on any path, the value is an unknown `Variant`, or the array comes from an unresolved/external call. |

The rule may report only one `VBA249` finding for a single expression and
failure kind. Existing warning rules retain ownership of possible failures;
for example, `VBA227` continues to report an allocation or bound assumption
that is not proven safe, while `VBA249` is reserved for a deterministic
failure. Compile-equivalent rules such as `VBA228` and `VBA229` retain their
existing ownership and preflight behavior.

## Fact and path model

The rule consumes facts after the shared procedure CFG/dataflow analysis reaches
a fixed point:

- `constexpr.Result` supplies typed `Known`, `Unknown`, and `Invalid` values;
- type facts identify scalar, array, object, and `Variant` shapes;
- CFG reachability excludes impossible constant branches;
- joins retain a known value only when all relevant incoming paths agree;
- definite assignment and existing allocation/bound facts are reused; and
- exceptional, recovered, unresolved, and uncertain edges preserve the
  conservative pre-statement state.

An assignment, external/member call, unresolved procedure, or possible
`ByRef` mutation invalidates a prior fact when the shared transfer model cannot
prove that the value remains unchanged. Unknown values are not errors. A
runtime-error finding therefore requires a concrete value and a concrete
failure contract at the expression, not merely a type that could fail at
runtime.

For a statement such as:

```vb
x = 10 / denominator
```

the rule is silent unless the dataflow state proves `denominator = 0` on every
relevant path. The equivalent literal expression `x = 10 / 0` is deterministic
and is reported. A branch that assigns zero on one path and a nonzero or
unknown value on another path is silent at the use site.

## Array rules

Known fixed-array bounds, `Option Base`, and `ReDim`/`Erase` state are taken
from the shared array model. A dynamic array begins unallocated, `ReDim` can
establish an allocation and bounds, and `Erase` returns a dynamic array to the
unallocated state. Branch joins and uncertain exceptional paths retain only
facts that are proven on all paths. Unknown rank, unknown bounds, external
array returns, and uncertain `Variant` values are fail-open.

`VBA249` reports only the access expression that is proven to fail. It does
not infer an out-of-bounds error from an unknown zero-based/one-based origin,
and it does not treat a scalar or an unknown `Variant` as an unallocated array.
Possible array misuse remains under `VBA227`, with duplicate findings at the
same expression suppressed according to the diagnostic ownership policy.

## Unsupported or deferred candidates

The following issue candidates require a separately specified runtime surface
and are not `VBA249` findings until their facts are equally deterministic:

- API-specific constant argument domains without a stable local contract;
- scalar operands to object-only operators such as `Is`;
- incompatible scalar/array binary operators when VBA coercion is unresolved;
- `TypeOf ... Is` expressions that depend on an incomplete type/dispatch model;
- late-bound member availability; and
- unknown `Variant`, late-bound object, unresolved external type, or
  locale-dependent coercion cases generally.

No candidate is promoted from this list merely because one execution path can
fail. It requires a shared fact and a positive/adversarial fixture pair before
the rule contract is expanded.

## Suppression, preflight, and JSON

Use `xlflow:disable-line VBA249`, `xlflow:disable-next-line VBA249`, or
`[analyze].disabled_rules = ["VBA249"]` for an intentional local exception.
Suppression removes the analyzer finding from the normal projection; it does
not change the underlying fact or make the source compile-equivalent. Since
`preflight_blocking` is false, an unsuppressed finding remains visible in
`analyze`, `check`, and LSP output but does not stop a workbook-bound command's
source-preflight gate.

The additive context has this shape:

```json
{
  "code": "VBA249",
  "severity": "error",
  "runtime_error": {
    "kind": "division_by_zero"
  }
}
```

Consumers must use the registry's `evidence_class` to distinguish this
`runtime-error` finding from a compile-equivalent `error`; severity alone is
not a preflight policy signal.

## Verification requirements

Every supported kind requires positive and adversarial fixtures. The focused
matrix must include:

- literals, `Const`/`Enum` values, typed suffixes, and equivalent parenthesized
  expressions;
- a value that is zero on one branch but nonzero or unknown on another;
- unknown `Variant`, late-bound, external, `ByRef`, and locale-dependent cases;
- unreachable constant branches and recovered/uncertain CFG edges;
- conversion boundary values immediately inside and outside each supported
  target range;
- fixed and dynamic arrays across allocation, `ReDim`, `Erase`, `Option Base`,
  multidimensional bounds, and branch joins; and
- duplicate ownership cases where `VBA227` or a compile-equivalent rule must
  remain the only finding.

Batch and real-time projections must agree on the same source revision and
must preserve deterministic ordering and source ranges. Runtime-only cases do
not require VBE-oracle promotion. If a case is discovered to be a VBE compile
rejection, bind accepted and rejected oracle controls and move it to the
compile-equivalent contract instead of widening `VBA249`.

Real-world corpus changes are evidence, not automatic truth. Review each new
`VBA249` candidate as a true positive or false positive, add a focused
regression for every confirmed false positive, and update snapshots only after
the review ledger and deterministic verify-only checks agree.

## Related

- Issue #596
- `docs/adr/ADR-0042-deterministic-runtime-error-diagnostics.md`
- `docs/specs/vba-constant-expression-evaluation.md`
- `docs/specs/vba-control-flow-graph.md`
- `docs/specs/vba-source-sink-dataflow.md`
- `docs/specs/array-lifecycle-safety.md`
- `docs/specs/cli-contract.md`
