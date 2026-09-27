# VBA IsMissing Diagnostics

<!-- xlflow-rule-contract: {"id":"VBA283","family":"analyze","category":"correctness","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_invalid_ismissing_usage","inline_suppressible":true,"preflight_blocking":false} -->

`VBA283` is an opt-in, non-blocking warning available in batch analysis and
realtime/LSP diagnostics. It is disabled by default and supports
`xlflow:disable-line` and `xlflow:disable-next-line` suppressions.

## Eligibility contract

The rule checks calls resolved as the VBA `IsMissing` intrinsic, either
unqualified or explicitly qualified with `VBA.`. Calls shadowed by a project
procedure or non-callable declaration, calls on other receivers, and
unresolved or ambiguous calls are not treated as the intrinsic.

The argument is valid only when it directly names a parameter of the
containing procedure and that parameter is declared `Optional`, has Variant
type (an omitted type is Variant), and is not an array or `ParamArray`. An
explicit default is allowed for an Optional Variant: `IsMissing` still detects
whether the caller supplied that argument. Microsoft documents this case in
its [named and optional arguments example](https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/understanding-named-arguments-and-optional-arguments),
which declares `Optional varCountry As Variant = "USA"` and checks it with
`IsMissing`.

Parentheses are transparent for this check: any number of parentheses around
an otherwise valid parameter reference are accepted. The analyzer treats
member access, a local or module variable, a different parameter, a literal,
an operator expression, or another call as an invalid argument. Findings
highlight the complete argument expression.

Calls with an argument count other than one, missing argument-expression
facts, recovered syntax, conditional compilation, or incomplete call
resolution fail open. The rule does not infer the result of `IsMissing` from
runtime values or inspect callers.

## Configuration

Enable the rule with:

```toml
[analyze]
detect_invalid_ismissing_usage = true
```

The equivalent stable rule ID is `VBA283` in
`[analyze].disabled_rules`. The default severity is `warning`; it does not
block source preflight.

## Verification contract

Focused coverage includes explicit and implicit Variant Optional parameters,
parenthesized references, non-Variant and non-Optional parameters, Variant
parameters with explicit defaults, arrays and `ParamArray`,
local/member/expression arguments, unrelated parameters, intrinsic
qualification, local shadowing, malformed arity, recovery, inline suppression,
default-off behavior, planner applicability, and batch/realtime parity.
