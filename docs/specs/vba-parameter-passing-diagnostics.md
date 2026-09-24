# VBA Parameter-Passing Diagnostics

<!-- xlflow-rule-contract: {"id":"VBA270","family":"analyze","category":"maintainability","default_severity":"information","scope":"procedure-local","realtime":true,"configuration_key":"detect_implicit_byref_parameters","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VBA271","family":"analyze","category":"reliability","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_assigned_byval_parameters","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VBA272","family":"analyze","category":"maintainability","default_severity":"information","scope":"procedure-local","realtime":true,"configuration_key":"detect_byref_parameters_can_be_byval","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VBA273","family":"analyze","category":"correctness","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_misleading_property_value_byref","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VBA274","family":"analyze","category":"maintainability","default_severity":"information","scope":"procedure-local","realtime":true,"configuration_key":"detect_redundant_byref_modifiers","inline_suppressible":true,"preflight_blocking":false} -->

`VBA270` through `VBA274` are opt-in, non-blocking diagnostics available in
batch analysis and realtime/LSP analysis. They are disabled by default and
support the normal `xlflow:disable-line` and `xlflow:disable-next-line`
suppression comments.

## Rule contracts

### VBA270 — implicit ByRef parameter

Reports an ordinary parameter whose declaration omits `ByVal` or `ByRef`.
VBA defaults that parameter to `ByRef`; the diagnostic asks projects that
prefer explicit API contracts to write the intended modifier. Property
Let/Set value parameters, event handlers, and `Implements` members are
excluded.

### VBA271 — assigned ByVal parameter

Reports a direct assignment, `Set` replacement, or other canonical IR write
to a parameter with effective `ByVal` semantics. The new value changes only
the local copy and cannot replace the caller's argument. The final value
parameter of `Property Let` and `Property Set` is treated as effectively
`ByVal`, regardless of the written modifier.

Changing members of an object passed `ByVal` is not a write to the parameter
binding and is not reported. Replacing that binding with `Set parameter = ...`
is reported.

### VBA272 — ByRef parameter can be ByVal

Reports an ordinary `ByRef` parameter when xlflow can prove that its binding
is never replaced. Eligibility includes ordinary `Public`, `Friend`, and
`Private` procedures, but excludes host event signatures, `Implements`
members, Property Let/Set value parameters, arrays, and `ParamArray`.

The proof starts with direct parameter writes and propagates through uniquely
resolved project-local calls. Positional and named arguments are mapped to
the callee signature. A write by a callee's effective `ByRef` parameter is a
write by the caller; a resolved `ByVal` parameter is not. Ambiguous,
unresolved, external, recovered, or conditional-compiled call boundaries are
treated as possibly writing and suppress the diagnostic.

The proof tracks replacement of the argument variable. Mutating a known
object's members through a parameter does not require the caller variable
itself to be `ByRef`. A member write on a known user-defined type does mutate
the caller-visible value and therefore prevents the recommendation. Unknown
composite types fail open, including when a nested member is passed through a
potentially writing `ByRef` call.

### VBA273 — misleading Property value ByRef

Reports an explicit `ByRef` modifier on the final value parameter of
`Property Let` or `Property Set`. VBA applies `ByVal` runtime semantics to
that parameter even when the declaration says `ByRef`; spelling it `ByVal`
makes the signature match its behavior. Index parameters before the final
value parameter retain their ordinary passing semantics and are not reported
under this rule.

### VBA274 — redundant explicit ByRef

Reports explicit `ByRef` on ordinary parameters for projects that prefer the
VBA default spelling. Property value parameters, event handlers, and
`Implements` members are excluded.

## Configuration conflict

`detect_implicit_byref_parameters` and
`detect_redundant_byref_modifiers` encode mutually exclusive style policies.
Configuration loading fails when both are `true`; neither is enabled by
default. The semantic rules `VBA271` through `VBA273` are independent of this
style choice.

## Parity mapping

| Inspection intent               | xlflow   | Coverage notes                                                                       |
| ------------------------------- | -------- | ------------------------------------------------------------------------------------ |
| Implicit ByRef modifier         | `VBA270` | Ordinary procedures; constrained signatures excluded.                                |
| Assigned ByVal parameter        | `VBA271` | Canonical IR writes, including effective Property value semantics.                   |
| Parameter can be ByVal          | `VBA272` | Direct plus project-local call-chain mutation summary; unknown boundaries fail open. |
| Misleading Property value ByRef | `VBA273` | Final Property Let/Set value parameter only.                                         |
| Redundant ByRef modifier        | `VBA274` | Mutually exclusive configuration with `VBA270`.                                      |

## Verification contract

Focused tests cover direct and transitive writes, read-only parameters,
unknown targets, named arguments, object-member mutation versus UDT-member
mutation and binding replacement, array exclusion, Property Let/Set final-parameter semantics,
host event and `Implements` exclusions, style configuration conflict, inline
suppression, default-off behavior, batch/realtime parity, and filesystem /
in-memory parity.

The Property Let/Set rule describes runtime passing behavior, not a VBE
compile rejection, so all five registry entries are
`compile_equivalent: false`. Compile-only oracle fixtures cannot establish
caller-visible mutation semantics; runtime language-observation evidence is
kept separate from compile-equivalent diagnostic binding.
