# ADR-0054: Two-Level Discarded Function Return Diagnostics

## Status

Accepted

## Context

VBA lets any `Function` or `Property Get` be invoked as a standalone
statement, in which case the returned value is produced and silently dropped.
That pattern is sometimes intentional (a procedure used only for side
effects) and sometimes a bug (a result the caller forgot to consume).
Rubberduck distinguishes the two situations with separate
`FunctionReturnValueDiscardedInspection` and
`FunctionReturnValueAlwaysDiscardedInspection` rules, and Issue #822 asks for
equivalent coverage in xlflow: a call-site diagnostic for each discarded
invocation, plus a declaration-level diagnostic when every known caller
discards the result.

xlflow already resolves call sites to project procedures in batch, realtime,
and LSP analysis, and the procedure IR distinguishes call statements from
call expressions nested inside other statements. The open design questions
are how to prove that a return value is actually dropped rather than
consumed, which visibility boundary makes an "all callers discard" claim
sound, and how conservative the rules must be about late-bound and dynamic
dispatch.

## Decision

Add two opt-in analyzer rules that share one discard predicate and one
resolution contract.

| Rule     | Contract                                                                                                                                               | Default          |
| -------- | ------------------------------------------------------------------------------------------------------------------------------------------------------ | ---------------- |
| `VBA257` | A standalone `call_statement` that invokes a uniquely resolved `Function` or `Property Get` and drops the result is reported as a warning.             | Disabled; opt-in |
| `VBA258` | A `Private`/`Friend` `Function` or `Property Get` whose every statically resolved call site discards the result is reported as an information finding. | Disabled; opt-in |

A call discards its result only when the call-site range equals the range of
a `call_statement`-kind statement. Nested call expressions inside arguments,
conditions, and assignment values always occupy a strictly smaller range, so
`x = Foo()`, `If Foo() Then`, `Bar(Foo())`, and `Debug.Print Foo()` can never
be reported as discards. Indexed assignment targets such as `Foo(x) = v`
retain assignment statement kind, and `RaiseEvent`, `ReDim`, and `New`
expressions are excluded before the range comparison. The predicate is shared
by both rules so batch and realtime surfaces classify the same call the same
way.

Only a uniquely resolved, project-local `Function` or `Property Get` produces
evidence. `Sub` and `Property Let`/`Set` callees return nothing and never
qualify. Ambiguous, unresolved, member, external, builtin-like, and dynamic
resolutions stay silent in both rules; a `With obj : .Foo : End With` member
call that cannot be resolved is not a discard finding for `VBA257` and
suppresses the same-named candidate for `VBA258`.

`VBA258` is batch-only because it needs the resolved project call set. A
candidate must be an explicit `Private` or `Friend` procedure: implicit
visibility is `Public` in every module kind, and `Public` members remain
externally callable, so an all-callers claim is unsound for them. Event
handlers, recovered procedures, `#If`-branched declarations, and
`Interface_Member` procedures in a module with `Implements` are excluded
because their callers are not statically enumerable.

`VBA258` requires at least one resolved call site and reports only when every
resolved call site discards. Any of the following suppresses the finding: a
resolved call site that consumes the result; a same-named member, ambiguous,
unresolved, incomplete, external, or dynamic call; a non-call identifier or
member reference matching the candidate name outside any call-site range
(which covers `x = Foo`, `AddressOf Foo`, and late-bound `o.Foo` reads, while
write-only accesses and assignment targets such as the function's own
return-slot writes are excluded); a statically folded `Application.Run`,
`Application.OnTime`/`OnKey`, or `CallByName` target naming the candidate;
and any dynamic-dispatch argument that cannot be folded, which suppresses the
entire rule for the run. A candidate with zero resolved call sites produces
no finding because unused-procedure diagnostics own that case.

Both rules are non-blocking, inline-suppressible, and configurable through
compatibility keys `detect_discarded_function_return` and
`detect_function_return_always_discarded`, with `[analyze].disabled_rules`
remaining authoritative. `VBA257` runs on the batch and realtime/LSP
surfaces; `VBA258` runs on batch `analyze`/`check` only. The corpus workspace
enables both flags so real-world evidence accumulates without changing
production defaults.

## Consequences

- Accidentally ignored results become visible at each call site, while a
  function used only for side effects is identified once at its declaration.
- Consumed results, subs, and unresolvable callees stay quiet because the
  discard predicate is a strict statement-range equality check and the
  resolution check requires a unique project match.
- `VBA258` can miss a finding when dynamic dispatch, a `With` member call, or
  a same-named shadow makes the caller set uncertain; that fail-open
  direction is intentional and prevents false all-discard claims.
- The two rules intentionally overlap: an all-discard project also produces
  `VBA257` findings at each site. Consumers that want only the declaration
  signal can enable `VBA258` alone.
- The shared predicate, registry metadata, config mapping, corpus workspace
  flags, generated references, and the public specification must evolve
  together so batch, realtime, and corpus surfaces never diverge on the same
  source revision.

## Alternatives Considered

1. **A single rule with a severity switch.** Rejected because the call-site
   and declaration-level findings have different scopes, surfaces, and
   suppression semantics; merging them would couple a procedure-local
   realtime rule to a project-wide batch computation.
2. **Report `Public` functions for `VBA258`.** Rejected because public
   members are externally callable (macros, `Application.Run`, other
   workbooks), so the project call set is provably incomplete.
3. **Treat every unresolved same-named call as a discard.** Rejected because
   member calls and late-bound receivers may consume the result; uncertainty
   suppresses rather than counts.
4. **Require VBE oracle evidence.** Rejected because both rules are syntactic
   and resolution based; they make no compile-behavior claim, so the oracle
   contract does not apply.
5. **Enable by default.** Rejected because discarded results are a common
   deliberate idiom in real-world VBA; the rules stay opt-in until corpus
   evidence shows the signal is precise enough for a default policy change.

## Evidence

- Issue #822 (sub-issue of #816) defines the call-site and declaration-level
  requirements, including consumed-context coverage and conservative
  treatment of public APIs.
- `internal/analyze/discarded_return.go` implements the shared discard
  predicate, the `VBA257` per-procedure pass, and the `VBA258` project pass;
  `internal/analyze/discarded_return_test.go` covers batch/realtime parity
  and the suppression contract.
- `internal/vba/calls.DynamicReferencesForIR` exposes dynamic-dispatch
  targets (Application.Run/OnTime/OnKey, CallByName) so `VBA258` does not
  reparse source.
- `docs/specs/vba-discarded-function-return-diagnostics.md` records the
  public rule contract; `internal/staticanalysis/rules/registry.json` carries
  the shared metadata consumed by config, docs, and surface gating.
- `docs/adr/ADR-0024-shared-static-analysis-rule-registry.md` and
  `docs/adr/ADR-0053-default-member-semantic-resolution.md` establish the
  registry, fail-open resolution, and opt-in policy patterns this change
  follows.

## Related

- Issue #816 (parent)
- Issue #822
- ADR-0024, ADR-0053
- `docs/specs/vba-discarded-function-return-diagnostics.md`
- `docs/specs/vba-call-graph-reachability.md`
