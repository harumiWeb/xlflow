# Option Base Inconsistency Diagnostics

<!-- xlflow-rule-contract: {"id":"VBA271","family":"analyze","category":"correctness","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_option_base_paramarray_inconsistency","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VBA272","family":"analyze","category":"correctness","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_option_base_array_inconsistency","inline_suppressible":true,"preflight_blocking":false} -->

`Option Base 1` changes the default lower bound of `Dim`/`ReDim` array
declarations, but it does not change the lower bound of arrays produced by
the type-library-qualified `VBA.Array` intrinsic or passed through a
`ParamArray` parameter: both remain zero-based. `VBA271` reports a
`ParamArray` parameter declaration inside a module that declares
`Option Base 1`, and `VBA272` reports a qualified `VBA.Array(...)` call in
the same module context. Both rules are opt-in and disabled by default.
Enable them with `[analyze].detect_option_base_paramarray_inconsistency =
true` and `[analyze].detect_option_base_array_inconsistency = true`, or
suppress a site with `xlflow:disable-line` / `xlflow:disable-next-line`.

## Eligibility contract

Both rules activate only when the module declares `Option Base 1`. A module
with `Option Base 0` or no `Option Base` statement produces no findings for
either rule. `Private Module` files and class modules follow the same
declaration check.

`VBA271` candidates are parameters declared with the `ParamArray` keyword.
The diagnostic anchors on the parameter identifier inside the enclosing
procedure's parameter list. `ParamArray` is always a `Variant` array with
lower bound 0 regardless of `Option Base`, so no resolution or dataflow is
required.

`VBA272` candidates are call sites whose base name is `Array` with an
explicit `VBA` receiver. The `VBA.` qualifier names the type library itself
and cannot be shadowed by lexical or project declarations, so every
`VBA.Array(...)` call under `Option Base 1` is reported without consulting
call resolution. An unqualified `Array(...)` call is never reported: per the
VBA language reference, the unqualified intrinsic honors `Option Base`, so
under `Option Base 1` it already returns a one-based array consistent with
the module's declared base. A call qualified by any other receiver
(`obj.Array(...)`, `Module.Array(...)`) may be a project or host member
returning an arbitrary lower bound and stays silent. Indexed assignment
targets such as `VBA.Array(0) = v` are call-shaped facts but are never
invocations and are not reported.

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

`VBA271` iterates the declared parameter list only. `VBA272` performs a
single pass over the procedure's already-materialized call sites and needs
no resolution pass because the `VBA.` qualifier is unshadowable. Neither
rule adds planner requirements, fixed-point work, or interprocedural
evaluation. Diagnostic ordering follows the analyzer's deterministic
source-order sort.

## Boundaries

Both rules are syntactic; they do not depend on VBE compile behavior, so no
VBE oracle evidence is required. The zero-based contract of `VBA.Array` and
`ParamArray`, and the `Option Base` sensitivity of the unqualified `Array`
intrinsic, are established VBA language semantics documented in the VBA
language reference (Array function, Option Base statement). These rules
cover the same inspection intent as the Rubberduck `Option Base`
inspections for `Array` and `ParamArray`, implemented against xlflow's
call-site facts and module `Option Base` model rather than Rubberduck's
parse tree.
