# Discarded Function Return Diagnostics

<!-- xlflow-rule-contract: {"id":"VBA257","family":"analyze","category":"correctness","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_discarded_function_return","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VBA258","family":"analyze","category":"maintainability","default_severity":"information","scope":"project-wide","realtime":false,"configuration_key":"detect_function_return_always_discarded","inline_suppressible":true,"preflight_blocking":false} -->

`VBA257` reports a call site that invokes a resolved `Function` or
`Property Get` as a standalone statement and drops the returned value.
`VBA258` reports a `Private`/`Friend` `Function` or `Property Get` whose
every statically resolved call site discards the result, so the procedure
effectively behaves like a `Sub`. Both rules are opt-in and disabled by
default. Enable them with `[analyze].detect_discarded_function_return = true`
and `[analyze].detect_function_return_always_discarded = true`, or suppress a
site with `xlflow:disable-line` / `xlflow:disable-next-line`.

## Discard contract

A call discards its result only when the call is the complete statement. The
predicate requires a `call_statement` whose range equals the call-site range:
`Foo`, `Foo x`, `Call Foo()`, `Foo()`, `Mod.Foo`, and `Me.Foo` as whole
statements all qualify, including single-line `If flag Then Foo` bodies and
colon-separated `Bar : Foo` statements. Calls nested inside an expression
never qualify: `x = Foo()`, `If Foo() Then ...`, `Bar(Foo())`,
`Debug.Print Foo()`, and `Set x = Foo()` all consume the result.

Call-shaped syntax that is not an invocation is excluded. Indexed assignment
targets such as `Foo(x) = v` keep assignment semantics rather than `Call`
semantics, `RaiseEvent` sites are events rather than function calls, and
`ReDim`/`New` expressions never reach the predicate. A bare `x = Foo`
without parentheses is a consumed reference, not a discarded call.

The grammar can split a paren-less member call whose member name is a
reserved keyword — for example `web_Http.Open MethodToName(x), url, async` —
into an `expression_statement` covering the member access plus a bogus
`call_statement` covering the argument list. A call statement that follows a
non-label sibling statement on the same line with only whitespace between
them matches this mis-split signature and is not a discard finding for either
rule; legitimate same-line statements are separated by a colon (which the
grammar folds into a `label_statement`) or by block keywords such as `Else`.

## Resolution contract

Only a uniquely resolved project `Function` or `Property Get` produces
evidence. Subs and Property Let/Set never qualify because they return no
value. Ambiguous, unresolved, member, external, builtin-like, and dynamic
resolutions stay silent for both rules; a `With obj : .Foo : End With` member
call that cannot be resolved to a specific project procedure is not a
discard finding.

## VBA258 eligibility and suppression

`VBA258` is a batch-only, project-wide check. A candidate must be an explicit
`Private` or `Friend` `Function`/`Property Get`: implicit visibility is
`Public` in every module kind, and `Public` members remain externally
callable, so they are never reported. Event handlers, recovered procedures,
procedures declared under `#If` conditional branches, and
`<Interface>_<Member>` procedures in a module that declares `<Interface>`
with `Implements` are excluded because their callers are not statically
enumerable; an unrelated underscored helper such as `Parse_Name` in the same
module keeps coverage. The prefix match uses the complete declared interface
name because interface names may contain underscores: `Implements I_Foo`
makes the `Bar` implementation `I_Foo_Bar`.

The all-discard claim requires a complete project view. `VBA258` stays
silent for the whole run when a `PathFilter` excludes modules or any file
carries a parser error or missing-node flag, because hidden source can
conceal a consuming caller or a dynamic-dispatch reference (dynamic
dispatch reaches `Private` procedures too).

A candidate is reported only when it has at least one resolved call site and
every resolved call site discards the result. Any of the following suppresses
the finding:

- a resolved call site that consumes the result;
- a member, ambiguous, unresolved, incomplete, external, or dynamic call
  whose callee name matches the candidate;
- a non-call identifier or member reference matching the candidate name,
  such as `x = Foo`, `AddressOf Foo`, a late-bound `o.Foo` read, or a
  by-reference argument like `RedirectInstance Init, VarPtr(Init), Me, c`.
  The reference still suppresses when it sits inside another call's range as
  an argument or receiver; only the callee token of a call-site itself is
  exempt, so a same-named argument such as the second `Foo` in `Foo Foo`
  still suppresses (write-only accesses and assignment targets, including
  the function's own return-slot writes, are excluded);
- a statically known dynamic-dispatch target naming the candidate through
  `Application.Run`/`OnTime`/`OnKey` or `CallByName`, including the
  receiverless `Run "name"` and With-block `.Run` forms once resolution
  has ruled out a same-named project procedure or non-callable local;
- any dynamic-dispatch argument that cannot be folded to a static string,
  which suppresses all VBA258 findings for the run.

A candidate with zero statically known call sites produces no finding; that
situation is owned by the unused-procedure diagnostic family rather than
VBA258.

## Boundaries

Both rules are syntactic and resolution based; they do not depend on VBE
compile behavior, so no VBE oracle evidence is required. VBA257 runs in batch
`analyze`/`check` and in the realtime/LSP path with parity. VBA258 runs only
in batch `analyze`/`check` because it requires the resolved project call
set.
