# ADR-0056: Option Base Inconsistency Diagnostics

## Status

Accepted

## Context

Issue #824 (sub-issue of #816) asks xlflow to cover the Rubberduck-style
`Option Base` inspections: `Option Base 1` changes the default lower bound
of `Dim`/`ReDim` array declarations, but it does not change the lower bound
of arrays produced by the `VBA.Array` intrinsic or received through a
`ParamArray` parameter — both stay zero-based. Developers who enable
`Option Base 1` and keep using those constructs hit off-by-one defects that
compile cleanly.

The open design questions are how to prove an unqualified `Array(...)` call
really means the VBA intrinsic rather than a user-defined or shadowed name,
and where indexed array syntax ends and an invocation begins when the
parser records both as call-shaped facts.

## Decision

Add two opt-in analyzer rules.

| Rule     | Contract                                                                                                                                                   | Default          |
| -------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------- |
| `VBA270` | An `Array(...)` call that resolves to the VBA intrinsic inside a module declaring `Option Base 1` is a warning anchored on the call site.                  | Disabled; opt-in |
| `VBA271` | A `ParamArray` parameter inside a module declaring `Option Base 1` is a warning anchored on the parameter, because the host always supplies lower bound 0. | Disabled; opt-in |

`Option Base 1` is the activation boundary: the effective module base comes
from the existing `arrayOptionBase` facts, so `Option Base 0` and files
without the directive stay silent in batch and realtime analysis alike.

`VBA270` trusts call resolution instead of name matching. Only
`ResolutionBuiltinLike` counts, and an explicit `VBA` receiver is required
for qualified calls, so `obj.Array(...)`, unresolved members, external and
dynamic dispatch, ambiguous shapes, and project procedures named `Array`
all stay silent. Two call-shaped facts are excluded before resolution is
consulted: indexed assignment targets (`Array(0) = v`, identified through
`procedureir.IsAssignmentTargetCall`) and unqualified references shadowed
by a lexical `Array` declaration. The shadow check is diagnostic-local
because the resolver deliberately omits array declarations from
`NonCallableNames` — an indexed read `arr(i)` is grammar-identical to a
call, and recording `arr` as non-callable would corrupt ordinary array
indexing. The rule therefore performs its own scope lookup over procedure
locals, parameters, and the module declaration projection rather than
changing resolver policy.

`VBA271` needs no resolution: `ParamArray` is a `Variant` array with lower
bound 0 regardless of `Option Base`, so the parameter flag alone is
evidence. Both rules ship opt-in; the corpus workspace enables them so
real-world evidence accumulates without changing production defaults.

## Consequences

- `Option Base 1` projects gain opt-in warnings for the two constructs that
  ignore the directive, closing the Rubberduck parity gap tracked by #824.
- A user-defined `Function Array` or a shadowing declaration named `Array`
  suppresses `VBA270` correctly, but an array variable named `Array` also
  means every unqualified `Array(...)` site in its scope is indexed access
  and unreported — the diagnostic cannot distinguish shadowed reads further
  without type-level proof, which is the intended fail-open direction.
- `VBA.Array` explicit qualification is always reported under
  `Option Base 1` because no lexical declaration can shadow a qualified
  library reference.
- No planner requirement, kernel, or fixed-point machinery is added; the
  rules iterate existing call-site and parameter facts, so analysis cost is
  negligible.

## Alternatives Considered

1. **Populate `NonCallableNames` with array declarations.** Rejected: the
   resolver intentionally treats `arr(i)` as call-shaped so indexed access
   keeps working; listing `Array` as non-callable would regress ordinary
   array indexing for every consumer of the shared overlay.
2. **Match call text `Array(` / `VBA.Array(` lexically.** Rejected: text
   matching cannot distinguish a project `Array` procedure or a shadowing
   variable from the intrinsic and would reintroduce the false positives
   the resolution boundary exists to prevent.
3. **One merged rule covering both constructs.** Rejected: the opt-in keys,
   severities, and remediation guidance differ (call-site review versus
   signature review), matching the precedent of `VBA257`/`VBA258`.
4. **Enable by default.** Rejected pending corpus evidence; the flags ship
   opt-in and the corpus workspace enables them for review.

## Evidence

- Issue #824 (sub-issue of #816) defines the required constructs and the
  `Option Base 1` activation boundary.
- `internal/analyze/option_base_diagnostics.go` implements both rules;
  `internal/analyze/option_base_diagnostics_test.go` covers intrinsic,
  qualified, shadowed, user-defined, bounded-array, and `ParamArray` cases.
- `internal/vba/procedureir/resolver.go` `declarationNames` documents why
  array declarations are excluded from `NonCallableNames`, which motivates
  the diagnostic-local shadow check.
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
