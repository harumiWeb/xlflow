# Option Base Inconsistency Diagnostics

<!-- xlflow-rule-contract: {"id":"VBA270","family":"analyze","category":"correctness","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_option_base_array_inconsistency","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VBA271","family":"analyze","category":"correctness","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_option_base_paramarray_inconsistency","inline_suppressible":true,"preflight_blocking":false} -->

`Option Base 1` changes the default lower bound of `Dim`/`ReDim` array
declarations, but it does not change the lower bound of arrays produced by
the `VBA.Array` intrinsic or passed through a `ParamArray` parameter: both
remain zero-based. `VBA270` reports an `Array(...)` call that resolves to
the VBA intrinsic inside a module that declares `Option Base 1`, and
`VBA271` reports a `ParamArray` parameter declaration in the same module
context. Both rules are opt-in and disabled by default. Enable them with
`[analyze].detect_option_base_array_inconsistency = true` and
`[analyze].detect_option_base_paramarray_inconsistency = true`, or suppress
a site with `xlflow:disable-line` / `xlflow:disable-next-line`.

## Eligibility contract

Both rules activate only when the module declares `Option Base 1`. A module
with `Option Base 0` or no `Option Base` statement produces no findings for
either rule. `Private Module` files and class modules follow the same
declaration check.

`VBA270` candidates are call sites whose base name is `Array`. An explicit
`VBA.Array(...)` receiver qualifies directly. An unqualified `Array(...)`
qualifies only when call resolution proves the VBA intrinsic
(`builtin-like` resolution): a project procedure named `Array` resolves to
a matched project call, and a lexical `Dim`/`Const`/parameter or
module-level declaration named `Array` shadows the intrinsic, so neither is
reported. Indexed assignment targets such as `Array(0) = v` are call-shaped
facts but bind to a lexical array variable and are never reported. Any
other qualified receiver (`obj.Array(...)`, `Module.Array(...)`) and every
ambiguous, member, external, unresolved, or dynamic resolution stays
silent.

`VBA271` candidates are parameters declared with the `ParamArray` keyword.
The diagnostic anchors on the parameter identifier inside the enclosing
procedure's parameter list. `ParamArray` is always a `Variant` array with
lower bound 0 regardless of `Option Base`, so no resolution or dataflow is
required.

Both rules skip a procedure whose symbol sits inside conditional-compilation
(`#If`/`#ElseIf`/`#Else`) branches or whose IR is marked recovered, because
the call/parameter projection for such procedures may be incomplete; this is
a deliberate fail-open boundary, not a coverage claim.

Explicitly bounded arrays (`Dim a(1 To 5)`, `ReDim a(1 To n)`) and ordinary
indexed access to any array are outside both rules' scope; `Option Base`
does not affect them and they produce no finding here.

## Surfaces and configuration

Both rules are opt-in, warning-level, high-precision, non-blocking,
inline-suppressible, procedure-local analyzer rules available in batch and
realtime/LSP analysis with parity. `information` is a supported severity
override for both.

## Precision and performance

Both rules fail open when resolution is incomplete or ambiguous. `VBA270`
performs a single pass over the procedure's already-materialized call
sites plus a lexical declaration lookup; `VBA271` iterates the declared
parameter list only. Neither rule adds planner requirements, fixed-point
work, or interprocedural evaluation. Diagnostic ordering follows the
analyzer's deterministic source-order sort.

## Boundaries

Both rules are syntactic and resolution based; they do not depend on VBE
compile behavior, so no VBE oracle evidence is required. The zero-based
contract of `VBA.Array` and `ParamArray` is established VBA language
semantics. These rules cover the same inspection intent as the Rubberduck
`Option Base` inspections for `Array` and `ParamArray`, implemented against
xlflow's resolver and module `Option Base` model rather than Rubberduck's
parse tree.
