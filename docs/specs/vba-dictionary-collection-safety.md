# VBA Dictionary and Collection Safety Analysis

## Scope

This specification defines the runtime-safety diagnostics for VBA
`Scripting.Dictionary` and `Collection` usage. The rules are inference and
portability checks, not compile-equivalent diagnostics, and never block source
preflight.

`VBA207` and `VBA213` remain opt-in for compatibility. `VBA230` through
`VBA235` are default-enabled warnings, are configurable through both their
legacy `[analyze]` keys and `disabled_rules`, and may be suppressed inline.

## Shared object-state analysis

The analyzer recognizes an early-bound `Scripting.Dictionary`, `As New`
construction, `New Scripting.Dictionary`, and late-bound
`CreateObject("Scripting.Dictionary")`. It also recognizes `Collection`
construction and direct local aliases of either object. Reassignment starts a
new identity and prevents facts about the former object from being applied to
the new one.

For each known identity, procedure CFG state may retain:

- whether the object is a Dictionary or Collection;
- fresh, empty, non-empty, or unknown content state;
- keys definitely added or removed;
- the known Dictionary `CompareMode` and explicit key-normalization forms; and
- aliases that still refer to the identity.

At a control-flow join, only facts true on every incoming path survive.
Conflicting assignments, escaping values, ambiguous calls, dynamic calls, and
unsupported operations degrade affected facts to unknown instead of proving a
warning-level violation.

Batch analysis may summarize uniquely resolved project-local helpers to a
fixed point. Supported summaries include Dictionary/Collection factories,
parameter-relative `Add`, `Remove`, `RemoveAll`, and `CompareMode` effects,
`.Keys`/`.Items` evaluation, and simple `.Exists` wrappers. Effects on alias
arguments are applied to their known identity. Recursive, ambiguous, external,
or dynamically resolved helpers do not establish definite content facts.
Real-time analysis uses only helpers uniquely resolvable in the active
document.

## Rule contracts

### `VBA207`: item and key guards

The rule tracks fresh objects, `Add`, `Remove`, `RemoveAll`, Dictionary
`.Exists` conditions, early exits from negated guards, and Collection `Count`
guards. Dictionary `.Item`, default-member access, or `.Remove` of a key known
to be absent is a `warning`. An unguarded operation whose key state is unknown
is `information`. A dominating `.Exists(key)` true branch is accepted.

For a Collection, a narrow compatibility probe is accepted only when `On Error
Resume Next` covers one access, optional `Err.Number` inspection and
`Err.Clear`, and is immediately restored with `On Error GoTo 0`. Broader
`Resume Next` scopes remain owned by `VB004` and `VBA214`; `VBA207` does not
duplicate them.

### `VBA213`: Dictionary key/value iteration

`For Each element In dictionary` enumerates keys. The rule reports at most one
finding per loop when the iterator is consumed as a Dictionary value or object.
Ordinary key use, iterator rebinding, and explicit `.Items` iteration are
accepted. Known aliases and uniquely summarized factories retain this existing
iteration contract.

### `VBA230`: CompareMode ordering

The rule reports assigning `.CompareMode` after the same Dictionary identity
has definitely received an entry. Assignment while fresh or after a definite
`RemoveAll` is accepted. Its legacy key is
`detect_dictionary_compare_mode_order`.

### `VBA231`: loop materialization

The rule reports `.Keys` or `.Items` evaluation in a loop body that may execute
more than once, including a uniquely resolved helper with that effect. Caching
the resulting array before the loop and using `For Each ... In dict.Keys` or
`dict.Items` as the loop's enumeration expression are accepted. Its legacy key
is `detect_dictionary_loop_materialization`.

### `VBA232`: key normalization

The rule reports when accesses derived from the same source key mix the raw
form with `LCase`/`LCase$`, `UCase`/`UCase$`, `Trim`/`Trim$`, or compositions of
those forms. Direct scalar aliases are followed. Case-only differences are
accepted when `vbTextCompare` is definite. User-defined normalizers are not
inferred. Its legacy key is `detect_dictionary_key_normalization`.

### `VBA233`: late-bound comparison constants

The rule reports `BinaryCompare`, `TextCompare`, or `DatabaseCompare` used for
the `CompareMode` of a late-bound Dictionary when that identifier is not
declared as a project-local `Const`. The built-in VBA constants
`vbBinaryCompare`, `vbTextCompare`, and `vbDatabaseCompare`, numeric values,
and explicit project constants are accepted. The TypeDB records the Scripting
enum names so the analyzer can identify this portability risk without treating
them as replacements for the VBA constants. Its legacy key is
`detect_late_bound_dictionary_constants`.

### `VBA234`: Collection mutation during iteration

The rule reports `Add` or `Remove` of the Collection identity currently used by
an enclosing `For Each`, including a uniquely summarized helper mutation. A
different Collection identity is accepted. Its legacy key is
`detect_collection_iteration_mutation`.

### `VBA235`: Collection index origin

The rule reports `.Item(0)`, an unadjusted index from a proven
`0 To collection.Count - 1` loop, and an unadjusted zero-origin array index used
against a Collection. `1 To collection.Count`, reverse one-based deletion, and
an explicit `index + 1` conversion are accepted. Its legacy key is
`detect_collection_index_origin`.

## Severity and surfaces

All findings are non-blocking. `VBA230` through `VBA235` use `warning`.
`VBA207` supports `warning` for definite absence and `information` for unknown
existence. Human output renders both values; JSON preserves the lowercase
severity string, and LSP maps `information` to
`DiagnosticSeverity.Information`.

The new rules run in batch and real-time analysis, subject to the helper-summary
scope described above. Inline suppression uses the normal
`xlflow:disable-line` and `xlflow:disable-next-line` contracts.

## Diagnostic ownership

- Broad `On Error Resume Next` remains owned by `VB004` and `VBA214`.
- Direct Dictionary iteration confusion remains owned by `VBA213`.
- Dictionary/Collection content guards remain owned by `VBA207`.
- Compile-time undeclared-identifier findings remain separate from the
  late-binding portability warning in `VBA233`.
