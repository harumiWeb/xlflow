# Unavailable WorksheetFunction Member Diagnostics

This specification defines `VBA252` for issue #690. It reports a member call
that is resolved against the complete generated Excel `WorksheetFunction`
type but is not present in that type's member set. The rule makes an invalid
object-model member visible before runtime without maintaining a hand-written
list of worksheet functions.

<!-- xlflow-rule-contract: {"id":"VBA252","family":"analyze","category":"correctness","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_unavailable_worksheet_function_members","inline_suppressible":true,"preflight_blocking":false} -->

## Public contract

`VBA252` is a default-enabled, warning-level, high-precision,
procedure-local analyzer rule in the correctness category. It is available in
batch and realtime/LSP analysis, is inline-suppressible, and does not block
source preflight. Its evidence class is inference: the resolved receiver type
and absent member are known from the generated TypeLib database, but the rule
does not claim that every Excel installation exposes the same version of the
object model.

The analyzer reports a call only when the receiver resolves to the typed
`Excel.WorksheetFunction` object and the generated TypeLib member set is
complete. Member names are compared case-insensitively using the normal
object-model resolver. The diagnostic range identifies the unavailable member
name at the call site. Qualified receivers such as
`Application.WorksheetFunction` and equivalent typed locals use the same
resolution path.

## Resolution and ownership boundaries

The generated TypeLib database is the source of truth; the rule does not ship
an allowlist of names such as `Abs` or `Concatenate`. If the TypeLib database
is missing, empty, malformed, partial, curated-only, stale, or otherwise
cannot prove a complete `Excel.WorksheetFunction` member set, the rule fails
open. A generated manifest is considered stale when its schema, generator, or
generator version metadata is incompatible with the running consumer.

Unresolved or ambiguous receivers, user-defined shadowing, `Object` and
`Variant` values, late-bound calls, and non-Excel members remain silent. The
rule owns unavailable typed `WorksheetFunction` members only; other Excel
member mismatches retain their existing diagnostic ownership. A resolved
unavailable member is not a source compile-equivalent error and therefore does
not affect preflight success.

## Configuration and suppression

The compatibility configuration key can disable the default-enabled rule:

```toml
[analyze]
detect_unavailable_worksheet_function_members = false
```

Projects may instead use the shared rule policy:

```toml
[analyze]
disabled_rules = ["VBA252"]
```

An intentional local exception can use `xlflow:disable-line VBA252` or
`xlflow:disable-next-line VBA252`. Suppression removes the finding but does not
change Excel's runtime behavior.

## Verification requirements

Focused coverage must include valid and unavailable members through direct,
qualified, typed-local, `With`, parenthesized, and multiline call forms. It
must also keep user-defined shadowing, unresolved/late-bound receivers,
incomplete TypeLib data, and non-`WorksheetFunction` members silent. Batch and
realtime projections must agree on finding identity, message, procedure, and
member range. Configuration disablement and both inline suppression forms must
be covered.

This rule is based on generated TypeLib metadata and does not require
VBE-oracle promotion or Excel COM execution. Corpus snapshot changes require
separate source review and are not part of this specification change.

## Related

- Issue #690
- [CLI contract](cli-contract.md)
- [Static-analysis diagnostic catalog](../../vitepress/reference/diagnostics.md)
