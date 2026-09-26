# ADR-0057: Class/Interface Public-API Hazard Diagnostics

## Status

Accepted

## Context

Issue #826 (sub-issue of #816) asks xlflow to cover the Rubberduck-style
class and interface inspections: VBA object modules carry naming
conventions with no syntactic guard rails — `<Interface>_<Member>` binds
an `Implements` implementation, `<Object>_<Event>` binds event
handlers, document modules own a predeclared default instance, and a
public Enum or write-only property in the wrong module shape produces
hazards that compile cleanly.

Two subtleties were settled during implementation review:

- The mandated lifecycle names `Class_Initialize` and `Class_Terminate`
  in class modules cannot satisfy the no-underscore convention because
  VBA reserves those exact identifiers, so flagging them is never
  actionable.
- Interface names may themselves contain underscores
  (`Implements I_Foo` binds `I_Foo_Bar`), so the implementation
  binding must match the complete declared interface name, not the first
  underscore segment.

## Decision

Add five opt-in analyzer rules in `internal/analyze/class_interface_hazards.go`.

| Rule     | Contract                                                                                                                                                      | Default          |
| -------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------- |
| `VBA273` | A public Sub/Function/Property in a class/form/document module whose name contains `_` collides with `<Interface>_<Member>` and `<Object>_<Event>` naming.    | Disabled; opt-in |
| `VBA274` | A non-Private `Enum` in a document module is duplicated into every runtime sheet copy, producing an ambiguous-name compile error.                             | Disabled; opt-in |
| `VBA275` | A public-facing `Property Let`/`Set` with no `Property Get` on the same name is a write-only API.                                                             | Disabled; opt-in |
| `VBA276` | A public member matching an `Implements` interface's `<Interface>_<Member>` binding, or an explicitly `Public` event handler, leaks to the default interface. | Disabled; opt-in |
| `VBA277` | A self-name reference inside a predeclared module (document/form kind or `VB_PredeclaredId = True` class) binds to the shared default instance, not `Me`.     | Disabled; opt-in |

Module kind is the activation boundary: VBA273 and VBA276 run only on
object-module kinds (class/form/document), VBA274 only on document
modules, VBA275 on any module with properties, and VBA277 only when the
module owns a predeclared default instance. All five read existing
procedure-symbol, declaration, and access-list metadata from the
procedure IR; none adds planner requirements, call resolution, or
fixed-point work, so they run identically in batch, realtime, and LSP
analysis.

Dedup boundaries are explicit: VBA273 skips recognized event handlers
(`IsEventHandler`), implemented-interface members (VBA276's
responsibility), and `Class_Initialize`/`Class_Terminate` in class
modules so each member produces at most one finding. VBA276 matches the
complete interface name plus `_`, so underscored interface names bind
correctly. VBA277 excludes type-position identifiers (`As` clauses,
`Implements` targets, `New` and `TypeOf ... Is` operands) through
the expression-parent chain.

## Consequences

- Projects gain five opt-in warnings covering the Rubberduck class /
  interface hazard set, closing the gap tracked by #826.
- `VBA276` produces at most one finding per member: an `I_Foo_Bar`
  public member colliding with `Implements I_Foo` is reported once on
  the implementation-binding contract, not again under the underscore
  rule.
- Predeclared classes whose style intentionally qualifies members through
  the class name (for example factory-style `stdDate.Create` patterns
  inside the class itself) will see VBA277 findings; they are opt-in and
  individually suppressible.
- No planner, kernel, or interprocedural machinery is added; each rule
  costs one symbol/declaration/access iteration per file or procedure.
- VBE oracle fixtures bind the compile-level facts only: the accepted
  cases confirm VBE tolerates the flagged shapes (so the findings are
  policy warnings, not compile errors) and the rejected interface
  collision case confirms the ambiguous-name mechanism behind VBA276.

## Alternatives Considered

1. **Match member names and declarations lexically.** Rejected:
   procedure-IR symbols already carry visibility, kind, event-handler,
   and Implements metadata, and text matching cannot distinguish
   `Class_Initialize`, Implements bindings, or type-position
   identifiers.
2. **One merged "class hygiene" rule.** Rejected: the five hazards have
   different activation boundaries, remediation guidance, and config
   keys, matching the VBA257/VBA258 precedent of separate opt-in rules.
3. **Resolve interface members across project modules to flag only
   actually-implemented members.** Rejected as unnecessary for the
   hazard: a public `<Interface>_<Member>`-shaped member is already the
   ambiguous-name / implementation-leak risk regardless of whether the
   interface declares that member.
4. **Enable by default.** Rejected pending corpus evidence; the rules
   ship opt-in and the corpus workspace enables them for review.

## Evidence

- Issue #826 (sub-issue of #816) defines the required hazard set and the
  opt-in/configurable policy.
- `internal/analyze/class_interface_hazards.go` implements all five
  rules; `internal/analyze/class_interface_hazards_test.go` covers each
  rule's activation boundary, exclusions, and default-off behavior.
- `docs/specs/vba-class-interface-diagnostics.md` records the public
  rule contract; `internal/staticanalysis/rules/registry.json` carries
  the shared metadata consumed by config, docs, and surface gating.
- VBE oracle fixtures `vba273-public-underscore-member`,
  `vba274-document-public-enum`, `vba275-write-only-property`,
  `vba276-public-event-handler`,
  `vba276-public-interface-member-collision`, and
  `vba277-predeclared-self-access` record the executed Excel
  16.0/17932/x64/ja-JP observations (five accepted, the interface
  collision rejected as ambiguous-name).
- Corpus review: 273 findings across nine third-party projects are
  committed as reviewed true-positives in
  `testdata/static-analysis-corpus/reviews/diagnostics.jsonl` with
  matching analyze snapshots.
- `docs/adr/ADR-0024-shared-static-analysis-rule-registry.md`,
  `docs/adr/ADR-0054-discarded-function-return-diagnostics.md`, and
  `docs/adr/ADR-0056-option-base-inconsistency-diagnostics.md`
  establish the registry, fail-open, and opt-in batch patterns this
  change follows.

## Related

- Issue #816 (parent)
- Issue #826
- ADR-0024, ADR-0054, ADR-0055, ADR-0056
- `docs/specs/vba-class-interface-diagnostics.md`
