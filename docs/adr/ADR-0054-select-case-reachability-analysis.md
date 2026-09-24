# ADR-0054: Select Case Reachability Analysis

## Status

Accepted

## Context

Issue #819 asks for a diagnostic that finds `Select Case` branches that can
never execute. VBA runs only the first matching `Case` clause in source order,
so an item shadowed by earlier items is dead code or, worse, evidence that a
broad earlier clause silently steals the intended branch. The VB063 parser
diagnostics already own `Case` syntax and ordering errors (`case_outside_select`,
`duplicate_case_else`, `case_after_else`); what is missing is the semantic
question of whether a syntactically valid item can ever match.

Two design pressures shape the rule. First, the analysis must never claim a
branch is unreachable when VBA's runtime coercion could still reach it — a
false "dead code" report invites users to delete live branches. Second, batch
`analyze` and realtime/LSP diagnostics must agree, so the rule needs to reuse
the shared constant evaluator and procedure constant environment rather than
re-derive values locally.

## Decision

Add `VBA259` as an opt-in, warning-level, non-blocking, inline-suppressible,
procedure-local analyzer rule (`detect_unreachable_select_case`), available in
batch and realtime analysis with `high` precision and `inference` evidence.

### Interval coverage model

Each `case_expression` item is classified into a value set: a singleton
point, a closed `To` range, or an `Is` comparison ray. Items accumulate into
a sorted disjoint union of the selector values already matched. A later item
is reported as `duplicate` when its effective set equals an earlier one,
`covered` when it is a subset of the union, `empty_range` when a `To` bound
inverts, and `impossible` when the selector's declared type domain cannot
produce a matching value. `Case Else` is reported as `else_covered` only when
the selector domain is finite (`Boolean`, `Byte`, `Integer`, `Long`, or a
statically known constant) and the union covers that domain.

Numeric intervals are compared as float64 closed intervals with
`math.Nextafter` for strict bounds; operands beyond float64's exact integer
range fail open. String intervals use code-unit order only under
`Option Compare Binary`; `Text` folds ASCII-only literals, and `Database`
strings stay unmodeled because its ordering is locale-dependent.

### Fail-open boundary

The rule reports nothing when a claim cannot be proven: recovered syntax or
conditional-compilation branches anywhere in the `Select`, unknown or
cross-type items (VBA coercion can still match `Case "5"` against a numeric
selector), `Date`/object/open-ended selector domains, and opaque expressions
all stay silent. Open `Variant` domains still allow `duplicate`/`covered`
proofs between items because first-match-wins is independent of the selector
type, but they never produce `impossible` or `else_covered`.

### Integration

`VBA259` runs as a planner projection (`featureSelectCase` gates it on
`Select` statements or `select case` text) so procedures without Select Case
pay nothing. It reuses the `procedureConstantEnvironment` extraction shared
with the runtime-error path, so batch and realtime findings are identical and
procedure-local names keep hiding same-named module constants.

The rule stays disabled by default. Parent issue #816 requires real-world
corpus evaluation before any default-enable decision.

## Consequences

- Users opt in to dead/shadowed `Select Case` detection that is precise enough
  to act on: every reported item is provably unreachable, not merely unusual.
- Batch and LSP surfaces share one constant model and cannot diverge on
  `Option Compare` or constant handling.
- Coverage is intentionally incomplete: late-bound coercions, `Like`-style
  comparisons, `Database` collation, and `Variant` runtime shapes remain
  silent even when a human could prove them unreachable.
- The `select_case_unreachable` finding context exposes `kind`,
  `item`, and `covered_by_line` so tooling does not parse message text.

## Alternatives Considered

1. **CFG/dataflow reachability for each clause.** Rejected: first-match-wins
   is a source-order property, not a CFG property; the interval model answers
   it directly and more cheaply while the CFG would add joins that cannot
   sharpen the proof.
2. **Reuse VB063 and extend it semantically.** Rejected: VB063 is a
   lint/parser diagnostic with syntax ownership; mixing semantic coverage
   into it would blur both rules' contracts and surfaces.
3. **Model `Option Compare Text`/`Database` ordering fully.** Rejected:
   locale-dependent collation cannot be reproduced deterministically; only
   ASCII folding under `Text` is modeled and everything else fails open.
4. **Default-enable with the other warning rules.** Rejected by the #816
   policy: new inference rules earn default-enablement only after corpus
   evaluation demonstrates the false-positive rate.

## Evidence

- Issue #819 defines the rule scope and opt-in requirement; #816 defines the
  corpus-evaluation gate for default-enablement.
- `docs/specs/vba-select-case-reachability.md` records the public contract,
  finding kinds, and fail-open boundary.
- `docs/specs/vba-constant-expression-evaluation.md` and ADR-0041 define the
  shared evaluator the rule reuses.
- `docs/adr/ADR-0046-procedure-applicability-planning.md` defines the
  feature/projection mechanism the rule plugs into.
- `internal/analyze/select_case_unreachable_test.go` covers each finding
  kind, `Option Compare` modes, finite-domain `Case Else`, and every
  fail-open boundary including batch/realtime parity.

## Related

- Issue #819, Issue #816
- ADR-0024 (rule registry), ADR-0041 (constant evaluation),
  ADR-0046 (procedure applicability planning)
- `docs/specs/cli-contract.md`
