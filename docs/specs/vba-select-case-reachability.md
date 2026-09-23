# VBA Select Case Reachability Diagnostics

<!-- xlflow-rule-contract: {"id":"VBA259","family":"analyze","category":"maintainability","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_unreachable_select_case","inline_suppressible":true,"preflight_blocking":false} -->

`VBA259` reports a `Case` item or `Case Else` branch that can never execute
because earlier `Case` items already match every selector value the branch
could handle. `Select Case` runs only the first matching clause in source
order, so an unreachable item is either dead code or a sign that the intended
branch is shadowed by an earlier, broader one.

## Modeled Case items

Each `case_expression` item in a clause is classified independently:

- A singleton value (`Case 5`, `Case "abc"`, `Case SomeConstant`).
- A closed range (`Case 1 To 10`, `Case "a" To "z"`).
- An `Is` comparison (`Case Is >= 5`, `Case Is <> 0`).

Items are evaluated through the shared constant-expression evaluator and the
procedure's constant environment, so module-level `Const` values participate
in coverage while procedure-local names hide same-named module constants.

## Finding kinds

The structured `select_case_unreachable` context carries a `kind` value:

- `duplicate` — the item's effective value set is identical to an earlier
  item's.
- `covered` — every value the item could match is already matched by earlier
  items.
- `empty_range` — a `To` range whose lower bound exceeds its upper bound can
  match nothing.
- `impossible` — the selector's declared type cannot produce a value the item
  matches (for example `Case 300` on a `Byte` selector, or `Case 2` on a
  `Boolean` selector).
- `else_covered` — earlier items already cover the selector's entire domain,
  so `Case Else` is unreachable. This requires a finite domain: `Boolean`,
  `Byte`, `Integer`, `Long`, or a statically known singleton selector.

## Selector domain

The selector's value domain comes from a statically known constant value
(singleton domain), a declared identifier's type (type domain), or an open
`Variant`/undeclared domain. Open domains still allow `duplicate` and
`covered` proofs between items but never `impossible` or `else_covered`.

`Option Compare` is honored for string items: `Binary` uses code-unit order,
`Text` folds ASCII-only literals, and `Database` ordering cannot be
reproduced deterministically, so string analysis under it stays silent.
Under `Text`, a constant string selector that is not ASCII-foldable also
stays silent, because real text comparison applies non-ASCII case mappings
(for example U+212A equals "k") that would disagree with the folded item
intervals.

## Fail-open contract

The rule emits nothing when it cannot prove a claim: recovered syntax,
conditional-compilation branches (`#If`/`#ElseIf`/`#Else`), unresolved or
cross-type items (VBA coerces types at runtime, so `Case "5"` on a numeric
selector is never flagged), unknown operands, `Date`/object selectors, and
values outside float64's exact integer range all skip silently. `VB063`
continues to own `Case`-clause syntax and ordering errors, including clauses
that follow `Case Else`.

## Surfaces and configuration

`VBA259` is an opt-in warning-level, non-blocking, high-precision analyzer
rule available in batch and realtime/LSP analysis. Enable it with
`detect_unreachable_select_case = true`, disable it per project with
`[analyze].disabled_rules`, or suppress a single item inline with
`xlflow:disable-line VBA259` / `xlflow:disable-next-line VBA259`. Item
findings point at the unreachable `case_expression` item; `duplicate` and
`covered` findings also record the earliest covering item's line as
`covered_by_line` in the context. `Case Else` findings point at the `Case
Else` keyword; because the proof usually spans several earlier items they
carry no single witness line.
