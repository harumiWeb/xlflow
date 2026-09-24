# ADR-0056: Conservative Parameter-Passing Diagnostics

## Status

Accepted

## Context

Issue #823 (sub-issue of #816) asks xlflow to cover five VBA parameter-passing
hazards and style choices. VBA defaults ordinary parameters to `ByRef`, while
the final value parameter of `Property Let` and `Property Set` has `ByVal`
runtime semantics even if its declaration says `ByRef`. A useful
"can be ByVal" diagnostic therefore needs mutation analysis rather than a
textual search.

Two requested style rules intentionally recommend opposite declarations:
one requires explicit `ByRef`, while the other removes redundant explicit
`ByRef`. Enabling both would create contradictory project policy.

## Decision

Add five opt-in, procedure-local, realtime-capable diagnostics:

| Rule     | Contract                                                                                                                       | Severity    |
| -------- | ------------------------------------------------------------------------------------------------------------------------------ | ----------- |
| `VBA270` | An ordinary parameter omits its passing modifier and therefore defaults to `ByRef`.                                            | information |
| `VBA271` | A parameter with effective `ByVal` semantics is reassigned.                                                                    | warning     |
| `VBA272` | An ordinary `ByRef` parameter is never directly reassigned or passed through a potentially writing `ByRef` boundary.           | information |
| `VBA273` | The final value parameter of `Property Let` or `Property Set` explicitly says `ByRef`, although VBA applies `ByVal` semantics. | warning     |
| `VBA274` | An ordinary parameter explicitly says `ByRef`, repeating VBA's default.                                                        | information |

All five rules are disabled by default. Configuration rejects enabling
`VBA270` and `VBA274` together, so a project must choose its declaration
style explicitly.

`VBA271` and `VBA272` use one project-local parameter-mutation summary.
Direct writes and read-writes mark a parameter as written. A resolved call
maps positional and named arguments to the callee signature and propagates a
write through effective `ByRef` parameters. Ambiguous, external, unresolved,
recovered, and conditional-compiled boundaries become `possibly written` and
suppress `VBA272`. Resolved `ByVal` and `ParamArray` arguments do not
propagate replacement of the caller's variable.

The analysis concerns replacement of the argument variable, not mutation of
an object it references. `target.Caption = ...` on a known object does not
write the `target` binding; `Set target = ...` does. User-defined types are
value-like: direct member writes and members passed through writing `ByRef`
calls count as caller-visible mutation. Unknown composite types fail open.
Arrays and `ParamArray` parameters are not eligible for `VBA272`, because VBA
does not permit typed arrays to be passed `ByVal`.

Host event signatures and `Implements` members are excluded from `VBA270`,
`VBA272`, and `VBA274` because their declaration shape is externally fixed.
Ordinary `Public`, `Friend`, and `Private` procedures remain eligible: these
rules describe their declared API contract rather than claiming a parameter
is unused by unknown callers.

## Consequences

- The semantic rules and style rules have separate IDs and configuration.
- Unknown calls reduce findings instead of creating false-positive
  "can be ByVal" claims.
- Mutation summaries are computed only when at least one of the five opt-in
  rules is enabled, preserving the default analysis cost.
- A project that wants explicit parameter contracts enables `VBA270`; a
  project that treats `ByRef` as idiomatic default syntax enables `VBA274`.
- Property value semantics are represented explicitly and are reused by
  call propagation as well as declaration diagnostics.

## Alternatives Considered

1. **Textually search for assignments.** Rejected because it misses writes
   propagated through local `ByRef` call chains and mishandles member writes.
2. **Treat every call argument as a write.** Rejected because resolved
   `ByVal` arguments are provably non-writing at the caller binding; unknown
   targets alone need the conservative boundary.
3. **Restrict `VBA272` to private procedures.** Rejected because changing a
   public/friend parameter from `ByRef` to `ByVal` is precisely an API-design
   recommendation. Event and interface contracts remain excluded.
4. **Allow both style flags and rely on per-parameter exclusivity.** Rejected
   because one project could then receive opposing recommendations across
   implicit and explicit declarations.
5. **Enable semantic rules by default.** Rejected pending corpus review; all
   five rules initially accumulate opt-in evidence.

## Evidence

- Issue #823 defines the five diagnostic families and requires actual
  mutation analysis plus conservative ByRef handling.
- `internal/analyze/parameter_passing.go` implements direct and propagated
  mutation summaries and the five projections.
- `internal/analyze/parameter_passing_test.go` covers writes, local call
  propagation, unknown calls, named arguments, object-member mutation,
  property value semantics, constrained signatures, suppression, realtime,
  and in-memory parity.
- `docs/specs/vba-parameter-passing-diagnostics.md` records the public
  contract, and the shared rule registry records surface metadata.

## Related

- Issue #816
- Issue #823
- ADR-0024, ADR-0046, ADR-0055
- `docs/specs/vba-parameter-passing-diagnostics.md`
