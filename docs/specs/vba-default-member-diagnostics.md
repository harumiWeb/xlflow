# VBA Default-Member and Bang-Notation Diagnostics

This specification defines the semantic default-member analysis requested by
issue #817. It covers implicit, indexed, recursive, unbound, and bang-notation
access while preserving the existing ownership of deterministic runtime error
and object-initialization diagnostics. ADR-0053 records the architectural
rationale.

<!-- xlflow-rule-contract: {"id":"VBA253","family":"analyze","category":"reliability","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_implicit_default_member_access","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VBA254","family":"analyze","category":"reliability","default_severity":"information","scope":"procedure-local","realtime":true,"configuration_key":"detect_unbound_default_member_access","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VBA255","family":"analyze","category":"maintainability","default_severity":"information","scope":"procedure-local","realtime":true,"configuration_key":"detect_bang_notation","inline_suppressible":true,"preflight_blocking":false} -->

## Public contract

The resolver consumes logical source statements, the revision-scoped document
expression/type and project-symbol view, retained `procedureir` member
operators, and the available TypeDB snapshot. It classifies one access once
and projects the result consistently through batch `analyze`, realtime
analysis, and the LSP workspace. It does not invoke Excel, VBE, or COM.

| Rule     | Severity and precision          | Reported access                                                                   | Default                                  |
| -------- | ------------------------------- | --------------------------------------------------------------------------------- | ---------------------------------------- |
| `VBA253` | `warning`, high precision       | Known implicit, indexed, or recursive default-member access                       | Disabled; opt-in after corpus evaluation |
| `VBA254` | `information`, medium precision | Unbound, late-bound, or incomplete default-member access                          | Disabled; opt-in                         |
| `VBA255` | `information`, high precision   | Bang notation (`receiver!name`) that uses stringly typed/default-member semantics | Disabled; opt-in                         |

All three rules are procedure-local, non-blocking, inline-suppressible, and
opt-in. Enable `VBA253` with:

```toml
[analyze]
detect_implicit_default_member_access = true
```

The other rules can be enabled with their respective keys:

```toml
[analyze]
detect_unbound_default_member_access = true
detect_bang_notation = true
```

The common rule policy is also supported:

```toml
[analyze]
disabled_rules = ["VBA253", "VBA254", "VBA255"]
```

An intentional local exception can use `xlflow:disable-line` or
`xlflow:disable-next-line` with the relevant rule ID. Suppression removes the
projected finding; it does not change the resolver facts or make an implicit
access explicit.

## Access classification

The additive `default_member` finding context has this shape:

```json
{
  "kind": "implicit",
  "binding": "known",
  "expected_context": "value",
  "member": "Value",
  "depth": 1
}
```

`kind` identifies the source/use shape (`implicit`, `indexed`, `recursive`, or
`bang`). `binding` is `known` only when the receiver and default member are
resolved from complete project or TypeDB facts; otherwise it is `unbound`.
For a `VBA249` projection, `binding` may additionally be `invalid` to mark a
complete negative proof; that value is not emitted by `VBA253` through
`VBA255`.
`expected_context` is `value` in the initial diagnostic surface. Object and
procedure uses remain unclassified until their call-context contract is backed
by focused VBE evidence.
`member` is the resolved default member when one is known, and `depth` is the
number of default-member hops in the resolved chain. The context is additive
to batch and internal realtime findings; the stable rule ID, severity, source
range, message, and normal finding fields remain unchanged. LSP diagnostics
preserve the same code, severity, message, and range.

The resolver distinguishes the following cases:

- **Implicit**: a value expression omits the member that VBA supplies from
  context, such as assigning a typed `Range` value where its default value
  member is required.
- **Indexed**: indexing/call syntax supplies arguments to a known default
  member, such as a collection or `Item` property. Explicit `.Item(...)` is
  not implicit access.
- **Recursive**: the result of one default member requires another default
  member before reaching the expected context. The chain is reported with its
  resolved depth.
- **Unbound**: the syntax may invoke a default member, but the receiver,
  member, or late-bound dispatch cannot be resolved with complete facts.
- **Bang**: `receiver!name` is retained as a distinct stringly typed access
  form. It may additionally be known, recursive, or unbound, but it is owned
  by `VBA255` when that rule is enabled.

Ordinary explicit member calls, explicit `.Value`/`.Item(...)`, comments,
strings, and names shadowed by a local declaration are not reported as
implicit access. A bang token is analyzed only as source syntax, not as proof
that a field or property named by its text exists.

## Resolution and fail-open policy

Generated TypeLib metadata may establish a default member and its signature
even when the wider database is incomplete. Completeness is required only for
negative proof. An exported project class-like module establishes its default
member when an exact `Attribute <member>.VB_UserMemId = 0` names a
value-producing Function or Property Get in that same revision. Commented
attributes, annotations, and naming conventions are not semantic evidence.
Missing, curated-only, or incomplete metadata does not prove that an external
type has no default member.

Uncertain facts never produce `VBA253` or a default-member-owned `VBA249`.
When `VBA254` is enabled, they may instead produce its explicitly unbound
advisory classification. Relevant uncertainty includes:

- TypeDB is missing or malformed, or a negative conclusion would rely on a
  stale, partial, or curated-only view;
- the procedure or expression is recovered/incomplete;
- the receiver is `Object`, `Variant`, late-bound, or dynamically dispatched;
- project symbols are ambiguous or a local declaration shadows the candidate;
- a call crosses an unresolved external or dynamic boundary; or
- the expected value context cannot be established.

Complete metadata can establish absence as negative evidence. This is the only
case in which an absent default member can support a deterministic failure
diagnostic. The resolver must not turn a TypeDB lookup miss into `VBA249` when
the metadata completeness contract is not satisfied.

## Expected contexts and diagnostic ownership

The initial semantic context is `value`, covering a simple ordinary assignment
or a directly indexed value expression. `Set` assignments and callable target
expressions remain silent; they are not reclassified by guessing from syntax.

`VBA253` reports a known valid implicit/default-member invocation. When the
default-member analysis is enabled and a complete model proves that the
expected context cannot be satisfied, ownership passes to `VBA249` and the
lower-confidence `VBA253` projection is suppressed for that expression. The
initial `VBA249` default-member runtime kinds are:

| Runtime kind              | Deterministic proof                                                                         |
| ------------------------- | ------------------------------------------------------------------------------------------- |
| `default_member_required` | A complete receiver type has no valid default member for the required value/object context. |
| `default_member_cycle`    | The default-member chain revisits a receiver/type before reaching a terminal result.        |

These are runtime-error evidence, not compile-equivalent errors, and remain
non-blocking for source preflight. An unresolved or possible 438-style failure
is `VBA254` at most when enabled; it is never `VBA249`.

`VBA202` retains exclusive ownership of deterministic error 91 when a receiver
is proven uninitialized or `Nothing`. If the receiver may be `Nothing`, the
default-member resolver fails open. A single expression must not receive both
the `VBA202` error-91 finding and a duplicate default-member runtime finding.

## Batch, realtime, and LSP behavior

All surfaces use the same resolver and revision-scoped semantic snapshot.
Findings preserve the source range of the omitted/default-member or bang access
and deterministic ordering. Realtime and LSP results must agree with batch for
the same complete source/project revision.

Fast LSP diagnostics do not publish project-negative conclusions for a pending,
partial, or unavailable workspace snapshot. Full LSP analysis may publish
`VBA253`/`VBA254`/`VBA255` after the complete overlay is available. An unsaved
sibling module or class is part of that overlay; stale workspace data is an
unknown state, not evidence that a member is absent.

## Verification requirements

Focused tests must cover:

- typed `Range` and collection defaults in value context;
- indexed defaults, nested and recursive chains, cycle detection, and
  explicit `.Value`/`.Item(...)` exclusions;
- generated TypeLib defaults, complete versus partial or curated-only TypeDB,
  local shadowing, `With`, and unsaved project overlays;
- `Object`, `Variant`, late-bound, unresolved, recovered, and ambiguous cases
  remaining fail-open;
- `receiver!name` versus string/comment text, known and unbound bang forms,
  suppression/configuration behavior, and diagnostic ownership with `VBA202`;
- deterministic `VBA249` default-member failures and duplicate suppression;
  and
- batch/realtime parity for additive context plus batch/realtime/LSP parity for
  code, severity, range, multiplicity, and ordering.

Corpus evaluation must run the new rules with their explicit opt-in settings
before any default policy is changed. Each candidate is reviewed against the
actual source evidence and complete diagnostic range. Confirmed false
positives require a focused regression and root-cause fix before the review
ledger or snapshot is updated. `allowed` evidence is separate from TP/FP
precision, and a start-only coordinate is not sufficient evidence for a ledger
entry.

If a proposed case is discovered to be a VBE compile rejection, bind accepted
and rejected local VBE controls and move it to the compile-equivalent contract;
do not widen `VBA249` from a runtime-only observation. Runtime-only default
member cases do not require Excel during ordinary analysis.

## Related

- Issue #816 (parent)
- Issue #817
- `docs/adr/ADR-0053-default-member-semantic-resolution.md`
- `docs/specs/vba-analysis-ir.md`
- `docs/specs/vba-resolution-diagnostics.md`
- `docs/specs/vba-runtime-error-diagnostics.md`
- `docs/specs/static-analysis-corpus.md`
