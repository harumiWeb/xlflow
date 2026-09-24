# Unused Declaration Diagnostics

<!-- xlflow-rule-contract: {"id":"VBA265","family":"analyze","category":"maintainability","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_unused_parameters","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VBA266","family":"analyze","category":"maintainability","default_severity":"warning","scope":"file-local","realtime":true,"configuration_key":"detect_unused_private_constants","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VBA267","family":"analyze","category":"maintainability","default_severity":"information","scope":"file-local","realtime":true,"configuration_key":"detect_unused_udt_members","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VBA268","family":"analyze","category":"reliability","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_never_assigned_variables","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VBA269","family":"analyze","category":"reliability","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_unassigned_variable_usage","inline_suppressible":true,"preflight_blocking":false} -->

`VBA265`–`VBA269` extend the unused-declaration analysis family. All five
rules are opt-in and disabled by default, run in batch `analyze`/`check` and
in the realtime/LSP path with parity, and support inline suppression. They
share the analyzer's conservative posture: unknown syntax, unresolved
resolution, or unmodeled control flow suppresses a finding rather than
producing one.

## VBA265 — unused procedure parameter

`VBA265` reports a parameter of a `Private` procedure that is never
referenced in the procedure body. Only `Private` procedures are candidates:
`Public` and `Friend` signatures are fixed by callers the analyzer cannot
enumerate. Event handlers, recovered or conditional-compiled procedures,
`Declare` statements, `<Interface>_<Member>` procedures in a module that
declares `Implements <Interface>` (a qualified `Implements Lib.IFace` target
still binds `IFace_<Member>` names), and procedures named by a string literal
anywhere in the project (the static surface of `Application.Run`, `OnTime`,
`OnAction`, `CallByName`, and similar dynamic dispatch) are excluded.

Signature-constrained event shapes are excluded even when the generic event
classifier does not recognize them: procedures whose name starts with a
`WithEvents` field name plus `_`, underscored names in document modules whose
signatures the host fixes (Access-style `secDetail_Format`, workbook and
worksheet object events), and incomplete `WithEvents` declarations all fail
open.

Use detection is lexical and intentionally permissive: a whole-word
occurrence of the parameter name anywhere in the procedure — including
inside comments and string literals — counts as a use and suppresses the
finding. Names that signal intent (`_`, empty, or names beginning with
`unused`/`ignore`) are never reported.

## VBA266 — unused private constant

`VBA266` reports a module-level `Private Const` that is never referenced.
A `Private Const` is file-local by language rules, so a single-module scan
is complete. References counted as use include procedure bodies, other
declaration initializers, and conditional-compilation directives; as with
`VBA265`, the scan is lexical and occurrences in comments or string literals
also mark the constant used.

When a procedure declares a local variable or parameter with the same name,
occurrences inside that procedure's body are attributed to the local and do
not count as constant references — except explicitly qualified
`Module.Const` occurrences, which still bind to the module-level constant
under a shadowing local. Recovered or conditional-compiled
declarations are skipped. Enum members are not `Const` declarations — they
may carry the constant flag in the IR, but `VBA266` only considers
declarations whose kind is `const`, so `Private Enum` members are never
candidates.

## VBA267 — unused user-defined type member

`VBA267` reports a member of a `Private Type` declaration that no resolvable
member expression in the module reads or writes. Private types are
file-local, so member usage is fully observable within the module.

Member usage resolves through:

- qualified member expressions (`p.X`, including nested chains such as
  `w.Inner.X` through member types that are themselves private UDTs);
- implicit member expressions inside `With` blocks (`.X`), including nested
  `With` statements and qualified heads inside them (`.Inner.X`);
- function return values whose declared return type is a private UDT in the
  same module; and
- `!` bang-notation member expressions.

Fail-open rules:

- an unresolved receiver marks the member name as used on every private UDT
  that declares it — this includes member expressions on a local, parameter,
  or function return slot that shadows a same-named module-level UDT
  variable with a non-UDT type;
- a type whose value flows through an unmodeled boundary — a `ByRef` or
  unresolved call argument at any expression depth inside the argument, a
  `Variant`/object/unknown assignment target, or an ambiguous statement
  (`Input #`, `Get #`, `LSet`, `Mid$` assignment) — is exempted entirely;
- member lines inside the `Type` block that cannot be modeled (conditional
  compilation, unusual declarations, duplicates) mark the whole type
  ineligible; and
- member names appearing inside string literals mark the member used, which
  covers `CallByName`-style dynamic access.

## VBA268/VBA269 — definite assignment

`VBA268` reports a procedure-local scalar variable that is read but never
assigned. `VBA269` reports a read that executes on a reachable control-flow
path where no assignment to the variable is guaranteed — the forward
definite-assignment fixpoint on the CFG, dual to the `VBA256` backward
liveness analysis.

Eligibility is deliberately narrow: only plain scalar locals (`Byte`,
`Integer`, `Long`, `LongLong`, `LongPtr`, `Single`, `Double`, `Currency`,
`Decimal`, `Date`, `Boolean`, `String`) qualify. Arrays, objects, Variants,
statics, constants, user-defined type variables, and the function return
slot are excluded because their initialization semantics or member-level
accesses are not provable at whole-variable granularity.

Assignment evidence includes direct writes, `ReDim`, `For` loop control
assignments, and arguments passed to a resolvable `ByRef` (or unresolved)
parameter — any call that cannot be resolved to a known `ByVal` signature is
treated as a potential definition of every argument. A potential ByRef
definition suppresses `VBA268` and joins the CFG's per-block assignment set
for `VBA269`, but the argument position itself is not counted as a proven
read, because initializer-style calls only write. The `Next <var>` occurrence
of a `For`/`For Each` loop is loop bookkeeping, not a value read — the loop
header itself performs the control-variable write — so it is excluded from
both read tracking and unassigned-read reporting.

`VBA269` measures definite assignment over normal control-flow edges only.
Exceptional edges (`On Error Resume Next` resume paths, error-handler jumps)
model writes that may be interrupted; treating them as ordinary predecessors
would empty every must-assignment set in procedures that use error handling
and flood the rule with false positives.

The whole procedure fails open when the CFG records reachable unknown-flow
sources or uncertain normal-class edges, when the IR marks the procedure or
the statement recovered, or when conditional-compilation branches cross the
modeled flow. `LSet`/`RSet`/`Mid$` assignment shapes exclude only the
affected variable. A variable that is only ever written on conditional paths
and then read unconditionally is a true positive under this contract: VBA
variables carry implicit default values, so the read is legal but relies on
the default.

## Verification contract

Focused tests cover each declaration category: unused and used parameters,
case-insensitive name matching, ignored-name conventions, public/friend,
event, `Implements`, and dynamic-entry exclusions for `VBA265`; constant
references in initializers and conditional compilation for `VBA266`; direct,
`With`-block, nested, and Variant-escape member resolution for `VBA267`;
never-assigned scalars, ByRef mutation, and straight-line/branching
definite-assignment reads for `VBA268`/`VBA269`. The rules do not depend on
VBE compile behavior, so no VBE oracle evidence is required.
