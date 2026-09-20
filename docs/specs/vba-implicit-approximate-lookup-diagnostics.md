# Implicit Approximate Lookup Diagnostics

This specification defines `VBA251` for issue #689. It reports Excel lookup
calls that omit the optional match-mode argument and therefore rely on Excel's
implicit approximate-match default. The rule makes lookup intent visible; it
does not prohibit intentional approximate matching or determine whether the
lookup data is correctly ordered at runtime.

<!-- xlflow-rule-contract: {"id":"VBA251","family":"analyze","category":"reliability","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_implicit_approximate_lookups","inline_suppressible":true,"preflight_blocking":false} -->

## Public contract

`VBA251` is a default-enabled, warning-level, high-precision, procedure-local
analyzer rule in the reliability category. It is available in batch and
realtime/LSP analysis, is inline-suppressible, and does not block source
preflight. Its evidence class is inference: the source omission is certain,
while the runtime result of the lookup is not. The rule uses the normal finding
envelope and does not add a CLI command, output-envelope version, or LSP
capability.

The analyzer reports a resolved Excel object-model call when all of the
following are true:

| API       | Omitted argument                 | Default behavior being made implicit |
| --------- | -------------------------------- | ------------------------------------ |
| `Match`   | third argument (`match_type`)    | approximate matching                 |
| `VLookup` | fourth argument (`range_lookup`) | approximate matching                 |
| `HLookup` | fourth argument (`range_lookup`) | approximate matching                 |

The supported calls include the typed `Application` and `WorksheetFunction`
receivers exposed by the Excel object model. The finding range covers the call
site and the message identifies the API and recommends making the matching mode
explicit.

Any explicit value in the mode position is accepted, including `0`/`False`,
`1`/`True`, a variable, a constant, or another expression. Named arguments,
mixed positional/named arguments, and multiline calls use the same positional
binding and are explicit when the mode parameter is supplied. The rule does not
judge whether an explicit approximate mode is correct for the workbook's data.

## Resolution and ownership boundaries

The rule requires typed receiver/member resolution before classifying a call.
User-defined procedures with the same name, unresolved or ambiguous calls,
late-bound `Object`/`Variant` receivers, and non-Excel members remain silent.
Name-shaped text alone is never enough to produce a finding.

`Lookup` is outside this contract because its optional result-vector argument is
not a match-mode switch. `XLookup` is also outside this contract because its
default match mode is exact. Other Excel APIs require a separate contract when
their omitted argument changes matching semantics.

This rule owns source-visible omission only. It does not check sort order,
duplicate keys, wildcard behavior, or the result returned for unsorted data.
Those runtime/data-dependent concerns remain covered by the Excel object-model
reference in `internal/agentskill/templates/xlflow/references/object-model-traps.md`.

## Configuration and suppression

The compatibility configuration key can disable the default-enabled rule:

```toml
[analyze]
detect_implicit_approximate_lookups = false
```

Projects may instead use the shared rule policy:

```toml
[analyze]
disabled_rules = ["VBA251"]
```

An intentional local exception can use `xlflow:disable-line VBA251` or
`xlflow:disable-next-line VBA251`. Suppression removes the finding but does not
change the call's Excel semantics. Because `preflight_blocking` is false, an
unsuppressed warning does not stop a workbook command before Excel opens.

## Verification requirements

Focused coverage must include each supported API through both typed receiver
families, omitted and explicit mode arguments, named and mixed arguments,
multiline calls, nested calls, and stable call-site ranges. It must also keep
same-named user procedures, unresolved/late-bound receivers, `Lookup`, and
`XLookup` silent. Batch and realtime projections must agree on source revision,
finding identity, message, and source range.

This is source-visible runtime-safety guidance and does not require VBE-oracle
promotion or Excel COM execution. Corpus snapshot changes require separate
source review and are not part of this specification change.

## Related

- Issue #689
- `internal/agentskill/templates/xlflow/references/object-model-traps.md`
- `docs/specs/cli-contract.md`
