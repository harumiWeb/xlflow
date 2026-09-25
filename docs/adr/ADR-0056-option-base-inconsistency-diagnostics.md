# ADR-0056: Option Base Inconsistency Diagnostics

## Status

Accepted

## Context

Issue #824 (sub-issue of #816) asks xlflow to cover the Rubberduck-style
`Option Base` inspections: `Option Base 1` changes the default lower bound
of `Dim`/`ReDim` array declarations, but it does not change the lower bound
of arrays produced by the type-library-qualified `VBA.Array` intrinsic or
received through a `ParamArray` parameter — both stay zero-based.
Developers who enable `Option Base 1` and keep using those constructs hit
off-by-one defects that compile cleanly.

A subtlety settled during review: the _unqualified_ `Array(...)` intrinsic
is not zero-based. Per the VBA language reference, its lower bound follows
the module's `Option Base`; only the form qualified with the type-library
name (`VBA.Array`) is unaffected. Flagging unqualified calls under
`Option Base 1` would therefore be a false positive, not an inconsistency.

The remaining design question is where indexed array syntax ends and an
invocation begins when the parser records both as call-shaped facts.

## Decision

Add two opt-in analyzer rules.

| Rule     | Contract                                                                                                                                                     | Default          |
| -------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------ | ---------------- |
| `VBA271` | A `ParamArray` parameter inside a module declaring `Option Base 1` is a warning anchored on the parameter, because the host always supplies lower bound 0.   | Disabled; opt-in |
| `VBA272` | A qualified `VBA.Array(...)` call inside a module declaring `Option Base 1` is a warning anchored on the call site, because it always returns lower bound 0. | Disabled; opt-in |

`Option Base 1` is the activation boundary: the effective module base comes
from the existing `arrayOptionBase` facts, so `Option Base 0` and files
without the directive stay silent in batch and realtime analysis alike.

`VBA272` requires the explicit `VBA` receiver and nothing else: the `VBA.`
qualifier names the type library itself and cannot be shadowed by lexical
or project declarations, so no call resolution is consulted. Unqualified
`Array(...)` is skipped entirely because it honors `Option Base`;
non-`VBA` receivers (`obj.Array(...)`) may be arbitrary members and stay
silent; indexed assignment targets (`VBA.Array(0) = v`, identified through
`procedureir.IsAssignmentTargetCall`) are call-shaped facts, not
invocations.

`VBA271` needs no resolution either: `ParamArray` is a `Variant` array with
lower bound 0 regardless of `Option Base`, so the parameter flag alone is
evidence. Both rules ship opt-in; the corpus workspace enables them so
real-world evidence accumulates without changing production defaults.

## Consequences

- `Option Base 1` projects gain opt-in warnings for the two constructs that
  ignore the directive, closing the Rubberduck parity gap tracked by #824.
- Unqualified `Array(...)` under `Option Base 1` is deliberately unreported:
  it produces a one-based array that already agrees with the module base.
  Callers migrating to `VBA.Array` for clarity inherit the warning and the
  documented remediation.
- `VBA.Array` explicit qualification is always reported under
  `Option Base 1` because no lexical declaration can shadow a qualified
  library reference.
- No planner requirement, kernel, fixed-point machinery, or call resolution
  is added; the rules iterate existing call-site and parameter facts, so
  analysis cost is negligible.
- `Parameter.NameRange` was added to procedure IR so `VBA271` anchors on the
  parameter identifier; the field is deep-copied and rebased on every
  incremental artifact-reuse path (`procedureir.Clone`/`RebaseProcedure`,
  `cfg.Clone`/`RebaseGraph`).

## Alternatives Considered

1. **Report unqualified `Array(...)` when it resolves to the VBA
   intrinsic.** Rejected during review: the unqualified intrinsic honors
   `Option Base`, so under `Option Base 1` the call is consistent and the
   warning would be a false positive.
2. **Match call text `VBA.Array(` lexically.** Rejected: receiver/basename
   facts already distinguish `foo.Array` from `VBA.Array`, and text
   matching would also catch comments or longer names.
3. **One merged rule covering both constructs.** Rejected: the opt-in keys,
   severities, and remediation guidance differ (call-site review versus
   signature review), matching the precedent of `VBA257`/`VBA258`.
4. **Enable by default.** Rejected pending corpus evidence; the flags ship
   opt-in and the corpus workspace enables them for review.

## Evidence

- Issue #824 (sub-issue of #816) defines the required constructs and the
  `Option Base 1` activation boundary.
- VBA language reference, "Array function" and "Option Base statement":
  `Array` follows `Option Base` unless qualified with the type-library
  name; `ParamArray` is always zero-based.
- `internal/analyze/option_base_diagnostics.go` implements both rules;
  `internal/analyze/option_base_diagnostics_test.go` covers qualified,
  unqualified, user-defined, shadowed, bounded-array, `ParamArray`, and
  batch/realtime parity cases.
- `docs/specs/vba-option-base-diagnostics.md` records the public rule
  contract; `internal/staticanalysis/rules/registry.json` carries the
  shared metadata consumed by config, docs, and surface gating.
- `docs/adr/ADR-0024-shared-static-analysis-rule-registry.md` and
  `docs/adr/ADR-0054-discarded-function-return-diagnostics.md` establish
  the registry, fail-open resolution, and opt-in policy patterns this
  change follows.

## Related

- Issue #816 (parent)
- Issue #824
- ADR-0024, ADR-0054, ADR-0055
- `docs/specs/vba-option-base-diagnostics.md`
