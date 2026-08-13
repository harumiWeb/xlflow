# VBA Opaque Boolean Control Analysis

<!-- xlflow-rule-contract: {"id":"VBA248","family":"analyze","category":"maintainability","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_opaque_boolean_arguments","inline_suppressible":true,"preflight_blocking":false} -->

This specification defines the opt-in `VBA248` call-site diagnostic and the
additive procedure metrics used to measure Boolean control complexity. The two
surfaces are intentionally independent: a call-site finding is about how a
caller communicates intent, while procedure metrics describe the declaration
without turning that measurement into a diagnostic.

## Call-site diagnostic (`VBA248`)

`VBA248` is a warning-level, non-blocking, procedure-local rule available to
batch and realtime analysis. It is disabled by default and is enabled with:

```toml
[analyze]
detect_opaque_boolean_arguments = true
```

The rule examines calls whose root argument expressions are known source
literals. A positional argument is a Boolean literal when its normalized text
is exactly `True` or `False`, ignoring case and redundant outer parentheses.
Variables, constants, expressions, and unknown values are not literals.

The staged heuristic is:

1. Report one finding for a call containing at least two positional Boolean
   literals. This high-confidence case is reported even when the callee is
   unresolved, because the ambiguity is visible at the call site.
2. A call containing one positional Boolean literal is reported only when a
   uniquely resolved local signature has at least two optional `Boolean`
   parameters. This catches an omitted switch while avoiding conventional
   single flags such as `overwrite`.
3. Named argument values are excluded from the positional count. A named
   argument therefore suppresses the single-literal case; remaining multiple
   positional literals still meet the first condition.

The finding range covers the call expression. The message identifies the
callee and literal count, and the suggestion recommends named arguments,
an enum, or separate procedures. When a local signature is uniquely resolved,
the suggestion includes the corresponding parameter names. Resolution that is
ambiguous, external, member-bound, or otherwise incomplete does not prevent a
multiple-literal finding, but it does not supply parameter-name suggestions.

The additive JSON context is:

```json
{
  "opaque_boolean": {
    "positional_literal_count": 2,
    "named_argument_count": 0,
    "parameter_names": ["first", "second"],
    "optional_boolean_parameter_count": 3
  }
}
```

`parameter_names` and `optional_boolean_parameter_count` are omitted when no
unique local signature supplies them. Inline suppression uses the normal
`xlflow:disable-line VBA248` and `xlflow:disable-next-line VBA248` forms; a
project can also suppress the rule with `[analyze].disabled_rules`.

## Declaration metrics

`xlflow metrics` adds the following non-threshold, additive integer fields to
each procedure entry. They are raw measurements and do not create `MX001`,
`VBA248`, or any other declaration-level finding:

| Field                                | Definition                                                                                                                                                                                                   |
| ------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `boolean_parameter_count`            | Declared parameters whose effective type is exactly `Boolean`.                                                                                                                                               |
| `optional_boolean_parameter_count`   | Boolean parameters declared with `Optional`.                                                                                                                                                                 |
| `vague_boolean_parameter_count`      | Boolean parameters whose complete, case-insensitive name is exactly `flag`, `mode`, or `option`. Prefixes and compound names are not included.                                                               |
| `boolean_control_branch_count`       | `If`/`ElseIf` statements whose complete condition is one Boolean parameter, optionally wrapped in `Not` and/or parentheses. Aliases, compound `And`/`Or` expressions, and interprocedural flow are excluded. |
| `boolean_controlled_statement_count` | The number of unique source-backed statements in the descendants of those directly controlled branches. Synthetic `do_condition` nodes and the branch statement itself are excluded.                         |

Metrics are collected for every procedure, regardless of `VBA248` enablement.
They use the existing metrics JSON schema version `1`; additive fields are
backward compatible. Consumers that need to compare versions must continue to
honor `metrics.schema_version`.

## Remediation guidance

For multiple independent behaviors, prefer an enum or separate procedures so
the API communicates the operation being requested. When compatibility requires
Boolean switches, named arguments make call sites self-documenting:

```vb
ProcessData first:=True, second:=False
```

The diagnostic does not infer business meaning, follow aliases, or judge a
parameter solely by a vague name. Those intentionally conservative limitations
keep the call-site signal high-confidence and leave broader API design review to
the raw declaration metrics.
