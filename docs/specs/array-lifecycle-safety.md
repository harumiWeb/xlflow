# Array Lifecycle and Dimension Safety

This specification defines the `VBA227` analyzer rule for issue #444 and the
array-operation validation boundary expanded by issue #593. It extends the
existing `ReDim Preserve`, array comparison, object-array `Set`, and
`Range.Value` checks through one conservative array-state model. The shared
shape contract and rationale are recorded in ADR-0040.

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

## Shared procedure result

For one procedure analysis revision, the array domain materializes one internal
`ArrayAnalysisResult` after the immutable procedure/module preparation has been
resolved. The result is worker-local, discarded with that revision, and
read-only after construction; it is never promoted to a project cache. Its
semantic payload covers the variable catalog, entry state, allocation and
shape/bound transitions, operation identity, branch refinements, ReDim facts,
and proven array runtime failures. Diagnostic projectors consume that payload
for `VBA227`, deterministic array `VBA249`, `VBA208`, `VBA209`, object-array
`VBA101`/`VBA102`, `VBA241`, and applicable `VBA226` without rebuilding the
catalog, constants, capacity guards, or operation lookup.

The procedure worker uses independent policy lanes over shared CFG scheduling.
The block-level lane preserves `VBA208`, `VBA249`, and object-array semantics;
the lifecycle lane retains `VBA227` physical source-line ordering, normal-edge
pruning, and reliable `On Error Resume Next` allocation behavior. `VBA226`
keeps its scalar/two-dimensional `Range.Value` lattice and may use one explicit
secondary walk. `VBA241` is a source/loop-region projection and is not counted
as another fixed-point walk. These policy choices are deliberate and do not
change diagnostic code, severity, message, range, ordering, suppression, or
Batch/realtime parity.

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
functions, and known array assignments establish allocation. An explicit
whole-array assignment such as `values() = Split(text, "|")` is an assignment,
not an indexed access. A whole-array argument such as `ConvertToJson(values(),
4)` is likewise not an indexed access at the call site; the callee owns any
element-access diagnostics. `ParamArray` values begin allocated even when the
caller supplies no arguments; their extent and bounds may remain unknown.
Unknown Variant and external values do not establish allocation.

A colon-separated declaration followed by an allocation, such as
`Dim values(): ReDim values(lower To upper)`, is one valid dynamic-array
allocation boundary. `VBA227` evaluates the `ReDim` before later source lines
and must not infer a fixed shape from the later `ReDim` bounds. A plain `ReDim`
therefore establishes the allocation before subsequent bounds and indexed
accesses are checked.

Known array-factory assignments such as `values() = Split(text, ",")` also
establish allocation when the VBA expression spans continuation lines. The
`VBA227` lifecycle pass joins those lines only for recognized `Array`, `Split`,
and `Filter` whole-array assignments before applying the allocation state; an
unrelated multiline call remains conservative.

The indexed-use scan shared by `VBA227` and `VBA249` considers only
unqualified local identifiers. A member call such as
`Application.OnTime(onTime, ...)` is not an access to a local array named
`OnTime`, even when a scalar with that name is declared in the same procedure.
Likewise, `driver_.TableToArray(...)` is a member call and not an indexed use
of the array-returning `TableToArray` procedure.

An `IsArray(variant)` true branch establishes array-ness for that Variant, so a
whole-array assignment from the guarded Variant can establish the target's
allocation. The false branch remains unknown until a recognized array factory
or other allocation operation establishes its state.

A `VarType` comparison (the VBA intrinsic) or the analyzer-supported
`VarTypeOf` guard name (which is not resolved as a project-defined wrapper)
with `(vbArray Or vbByte)` provides the same array proof for a guarded Variant
Byte-array assignment. Variant results
from a binary stream `Read(-1)` expression and the `vbNullString`-to-Byte-array
idiom are recognized as Byte-array transfers. The latter is a known empty
array: `LBound` / `UBound` queries are valid, but element access remains a
`VBA227` finding. A private `ByRef` output helper that fills the output on each
accepted branch and routes rejected inputs through a project-local procedure
with no normal exit is summarized at its normal exits; the rejecting branch
does not poison the allocation proof.

The same allocation-probe contract also applies when the positive length is
first assigned to a scalar local and that local is compared with zero or a
positive threshold. The proof remains path-sensitive; unrelated scalar
assignments do not establish allocation.

At a join, allocation and dimensions are retained only when all incoming paths
agree. Exceptional and uncertain CFG edges use the pre-statement state. This
means an indexed access after an allocation on only one branch remains a
warning.
Within a multiline CFG block, the `VBA227` lifecycle pass evaluates physical
source lines in order so an allocation earlier in the block is visible to a
later bound query or indexed access. This ordering refinement is local to
`VBA227`; `VBA208` and `VBA249` retain their existing CFG-block evaluation
semantics. On `On Error Resume Next` exceptional edges, a deterministic plain
`ReDim` or recognized array-factory assignment retains its established
allocation state; `ReDim Preserve` remains conservative when its prior
allocation or shape is unknown.
The rule also recognizes the narrow growable-buffer idiom that probes
`UBound(target) + 1` under `On Error Resume Next`, clears `Err`, restores error
handling, and conditionally performs `ReDim Preserve target(...)` before a
bounded write loop. The probe and that loop's indexed writes are treated as
covered by the fallback allocation; an unrelated `Resume Next` bounds query or
indexed access remains conservative.

The operation-facing shape is `scalar`, `fixed-array(rank)`,
`dynamic-array(rank/unknown)`, `Variant`, or `unknown`. Declaration metadata,
procedure signatures, array-return summaries, `Array(...)`, and `ParamArray`
are normalized into this shape before operation checks run. Assignments,
`ByRef` calls, and branch joins retain a shape only when it is proven on every
incoming path. For a unique project-local `Private` procedure, or a procedure
in an `Option Private Module`, a direct array argument may seed a `ByRef` array
parameter only when every observed call site passes an allocated array. These
entry facts are solved across the same restricted helper chain. A module-level
array-returning expression passed directly as a `ByRef` array argument may
also seed the parameter from its allocated return summary; unknown or
conditional returns remain unknown. Nested `ByRef` calls on one physical line
are handled only when their source ranges establish a safe evaluation order;
unrelated calls on the same line remain conservative. A module-level array may
also carry allocation from a dominating, unique project-local
caller into every resolved same-module `Private` helper entry. Class and form
procedures also inherit arrays proven by `UserForm_Initialize` /
`Class_Initialize`; cross-module, ambiguous, dynamic, unresolved, and otherwise
unproven callers remain unknown. An inline conditional setup call is not
sufficient. In a class module, a proven allocation performed by a
project-local configuration helper through a `ByRef` array parameter may also
be carried through a rejecting role guard or a matching role branch. A guard
that validates a configured collection role or kind—such as
`IsGenericCollectionRole`, `IsDictionaryCollection`, `IsSetCollection`,
`IsPriorityQueueKind`, `IsSortedMapKind`, `IsSortedSetKind`, a recognized
collection-role constant, or
`mCollectionKind`—may establish the arrays belonging to that generic
collection configuration. A helper name or guard without a proven allocation
is not sufficient. Public, ambiguous, dynamic, and unresolved calls remain
unknown. An unresolved or external value, and a `Variant` whose array nature
is not proven, remains `unknown`.

For class modules, the recognized `Friend` / `Private` collection, data-row,
and aggregate-error storage members inherit the matching `Configure*`
allocation contract when they are reached through a configured receiver. This
models internal storage accessors and mutators that deliberately rely on their
caller having established the owning instance's role; it does not establish
allocation for unrelated helpers or for public, ambiguous, dynamic, or
unresolved calls.

A private module setup helper may also establish its direct module-array
allocations when it uses the narrow idempotent form `If ready Then Exit Sub`,
performs plain `ReDim` statements, and assigns the module-scoped Boolean guard
to `True` as its final executable statement. The guard must have no other
assignment in the module and the array must have no other whole-array write or
`Erase`; arbitrary readiness flags, `ReDim Preserve`, externally writable
guards, and resettable buffers remain unknown.

A private `ByRef` output helper may establish a conditional allocation when it
exits for a zero `Collection.Count` (or equivalent count) and performs a
`ReDim` from that count on the remaining normal path. The caller must retain
that condition: a positive count branch or a positive numeric `Select Case`
clause proves the output array allocated only within that branch. The helper
does not make the array unconditionally allocated, and unrelated count
expressions or arbitrary conditional `ReDim` statements do not establish this
fact. `Select Case` headers and their case bodies are evaluated through their
own CFG blocks so a case-local access receives the selected branch state.

When a private `ByRef` output helper has an unconditional allocation summary,
that allocation is also applied to a caller-local array argument. This allows a
second private `ByRef` helper to consume the locally allocated result without
losing its entry state. Module-array effects remain shadow-safe: a module
allocation is not transferred into a same-named local variable, and unresolved
or public helper calls remain unknown.

A private `ByRef` helper may also return an array and its successful element
length through separate `ByRef` parameters. The assignment
`length = UBound(values) - LBound(values) + 1` establishes a conditional
allocation contract only when the helper's zero-length branch is explicit and
the caller tests that paired length positively. A call that passes an
unallocated array to a guarded helper is ignored for entry-state evidence when
the helper's array use is unreachable under literal `False` or non-positive
arguments; this prevents an intentionally unused optional array from
invalidating a reachable positive-length proof.

A non-empty `String` assigned directly to a dynamic `Byte` array is recognized
as an allocation for a non-empty literal, or for a statically typed `String`
whose syntactic non-empty guard is visible. An unguarded String assignment
remains unknown because an empty string can still leave no usable element
bounds; arbitrary `Variant` and function-return assignments remain unknown.

A private `ByRef` array helper that calls itself recursively preserves the
allocation state established by a proven external entry call. The recursive
edge is not treated as an independent unknown entry; public or otherwise
unresolved callers remain conservative.

## Diagnostics and ownership

`VBA227` reports only possible or partially known cases of:

- `LBound` / `UBound` before allocation, invalid dimension numbers, and
  dimension or bound contradictions when the allocation or shape is not fully
  proven;
- indexed access before `ReDim` or after `Erase` when allocation is possible
  but not proven to fail on every relevant path;
- `ReDim` on fixed-size arrays or known scalar values;
- `Erase` on a target proven incompatible with an array or object reset;
- `For Each` over a source proven not to be an array or collection (the
  existing `VB023` control-variable contract remains separate);
- known subscript-count mismatches and inconsistent loop lower-bound
  assumptions.

Fully proven `array_unallocated` and `array_subscript_out_of_bounds` accesses
are owned only by `VBA249`; `VBA227` must not duplicate those findings at the
same expression.

Constant declaration bounds and constant `ReDim` bounds are checked according
to their operation contract. Compile-equivalent declaration/assignment cases
use canonical blocking rule contracts. A runtime array failure that the shared
constant, type, and CFG/dataflow facts prove on every relevant path is
projected by `VBA249` as a `runtime-error` severity `error`; a possible or
partially known failure remains a non-blocking `VBA227` warning and does not
become a preflight claim. Dynamic bounds and unknown values are not diagnosed.

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
For object comparisons, `Set object = Nothing` assignments, including inline
`Then Set` assignments, are not equality comparisons and are excluded; scalar
`object = Nothing` comparisons remain diagnostics.
VBA declaration type characters (`$`, `%`, `&`, `!`, `#`, `@`, and `^`) are
explicit types for the shared `VB005` / `VB019` declaration checks. A second
`Dim` or `ReDim` keyword after a declaration comma is parser recovery and is
reported as `VB014`; its recovered fragments are not treated as additional
declarators or array lifecycle statements.
The common model also supplies object-array element assignments without `Set`
to the existing `VBA101` / `VBA102` finding constructors. `VBA226` remains the
owner of `Range.Value` / `Value2` shape diagnostics; Range-origin values are
excluded from duplicate `VBA227` access and bound findings.

## Function and property summaries

Batch analysis may use a unique project-local `Function` or `Property Get`
summary when every observed normal return assignment returns an allocated
array with a consistent shape. Real-time analysis restricts this summary to
the active document; it does not resolve array-return summaries from another
module. Batch summaries are solved to a fixed point across unique helper chains,
so declaration order does not change a proven result. Recognized allocation
guards refine the normal branch, and a definitely failing constant `ReDim`
without local error handling is excluded from normal-return evidence. Mixed
return kinds, missing assignments, recursive or ambiguous chains, and external
calls remain unknown.

A unique project-local scalar `Function` or `Property Get` with one array or
`Variant` parameter may also be recognized as an allocation probe when its
normal return is exactly `UBound(parameter) - LBound(parameter) + 1` or
`UBound(parameter) + 1` and its error-recovery label returns zero. For a typed
VBA function, falling through from that recovery label to the implicit default
return of zero is equivalent to an explicit zero assignment. The probe's own
bound reads are covered by that recovery contract and are not reported as
unallocated-array findings. A direct call comparison such as
`CountBytes(values) > 0`, `CountBytes(values) >= 1`, `CountBytes(values) <> 0`,
or the false branch of `CountBytes(values) = 0` then proves `values` allocated
on that branch, including when the helper accepts the value through a Variant
parameter. The rule does not infer allocation from arbitrary Boolean helpers,
compound conditions, or the opposite branch.

## Boundaries

The rule does not open Excel or claim that `IsArray` proves a dimension count.
Its true branch establishes only array-ness for a Variant whole-array
assignment, while a positive recognized allocation-probe result establishes
allocation for the guarded value. It accepts literal numeric bounds and the
small shared integer-constant subset (Const/Enum references with arithmetic)
only when they are statically visible; dynamic expressions remain unknown. It
does not change the `Finding` JSON envelope or preflight blocking behavior.

## Related

- `docs/adr/ADR-0028-array-lifecycle-and-dimension-safety.md`
- `docs/adr/ADR-0040-array-operation-shape-validation.md`
- `docs/adr/ADR-0027-range-value-array-shape-safety.md`
- `docs/specs/range-value-array-shape-safety.md`
- `docs/specs/cli-contract.md`
- Issue #444
- Issue #593
