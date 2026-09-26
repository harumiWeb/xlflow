# ADR-0057: Class/Interface Public-API Hazard Diagnostics

## Status

Accepted

## Context

Issue #826 (sub-issue of #816) asks xlflow to cover the Rubberduck-style
class and interface inspections: VBA object modules carry naming
conventions with no syntactic guard rails - `<Interface>_<Member>` binds
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

Two further constraints arrived during post-merge review:

- The original implementation trusted the IR's `IsEventHandler` fact for
  UserForms, which treats almost any underscored form member as an event,
  and trusted the `Implements` name prefix alone, which flags helpers
  that share an interface name without implementing a real member. Both
  had to be tightened against verifiable evidence: designer control names
  for form events and the interface class's declared public members for
  bindings.
- Rule IDs VBA273-VBA277 were already assigned on main to the
  parameter-passing family (issue #823), so the hazard rules ship as
  VBA278-VBA282.

## Decision

Add five opt-in analyzer rules in `internal/analyze/class_interface_hazards.go`.

| Rule     | Contract                                                                                                                                                              | Default          |
| -------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------- |
| `VBA278` | A public Sub/Function/Property in a class/form/document module whose name contains `_` collides with `<Interface>_<Member>` and `<Object>_<Event>` naming.            | Disabled; opt-in |
| `VBA279` | A non-Private `Enum` in a document module is duplicated into every runtime sheet copy, producing an ambiguous-name compile error.                                     | Disabled; opt-in |
| `VBA280` | A public-facing `Property Let`/`Set` with no `Property Get` on the same name is a write-only API.                                                                     | Disabled; opt-in |
| `VBA281` | A public member matching a verified `Implements` interface's `<Interface>_<Member>` binding, or an explicitly `Public` event handler, leaks to the default interface. | Disabled; opt-in |
| `VBA282` | A self-name reference inside a predeclared module (document/form kind or `VB_PredeclaredId = True` class) binds to the shared default instance, not `Me`.             | Disabled; opt-in |

Module kind is the activation boundary: VBA278 and VBA281 run only on
object-module kinds (class/form/document), VBA279 only on document
modules, VBA280 on any module with properties, and VBA282 only when the
module owns a predeclared default instance. All five read existing
procedure-symbol, declaration, and access-list metadata from the
procedure IR; the only added index is a per-file-set map of class-module
public member names used to verify interface bindings, built once per
analysis context. No rule adds planner requirements, call resolution, or
fixed-point work, so they run identically in batch, realtime, and LSP
analysis.

Dedup boundaries are explicit: VBA278 skips recognized event handlers,
verified or unresolved interface bindings (VBA281's responsibility), and
`Class_Initialize`/`Class_Terminate` in class modules so each member
produces at most one finding. VBA281 requires the suffix after the
interface prefix to name a public member of a resolved, unambiguous
interface class; a resolved interface without the member leaves the name
to VBA278, while an unresolved or ambiguous interface fails open on both
rules. Form events are recognized only when their control prefix exists
in the designer control set when that artifact is available, or matches
the intrinsic `UserForm_*` event names when controls are unknown. VBA280
treats conditional or recovered accessors as unprovable and counts Friend
writers as public-facing. VBA282 excludes accesses resolved to local,
parameter, or module scopes and narrows the `TypeOf` exclusion to the
type operand so the left-hand expression still counts.

## Consequences

- Projects gain five opt-in warnings covering the Rubberduck class /
  interface hazard set, closing the gap tracked by #826.
- `VBA281` produces at most one finding per member: an `I_Foo_Bar`
  public member colliding with `Implements I_Foo` is reported once on
  the implementation-binding contract, not again under the underscore
  rule.
- Interface bindings are verified against the analyzed file set:
  single-file and realtime contexts where the interface class is absent
  fail open rather than claiming a collision.
- Predeclared classes whose style intentionally qualifies members through
  the class name (for example factory-style `stdDate.Create` patterns
  inside the class itself) will see VBA282 findings; they are opt-in and
  individually suppressible.
- VBE oracle fixtures bind the compile-level facts only: the accepted
  cases confirm VBE tolerates the flagged shapes (so the findings are
  policy warnings, not compile errors) and the rejected interface
  collision case confirms the ambiguous-name mechanism behind VBA281.

## Alternatives Considered

1. **Match member names and declarations lexically.** Rejected:
   procedure-IR symbols already carry visibility, kind, event-handler,
   and Implements metadata, and text matching cannot distinguish
   `Class_Initialize`, Implements bindings, or type-position
   identifiers.
2. **One merged "class hygiene" rule.** Rejected: the five hazards have
   different activation boundaries, remediation guidance, and config
   keys, matching the VBA257/VBA258 precedent of separate opt-in rules.
3. **Trust the `Implements` name prefix alone.** Rejected during
   review: an underscored helper sharing an interface name was reported
   as a binding collision and suppressed the underscore rule. The
   analysis context now indexes class-module public members so the
   binding is verified where the interface resolves and fails open
   where it does not.
4. **Trust the IR `IsEventHandler` fact for form members.** Rejected
   during review: the fact accepts nearly any underscored form name, so
   public helpers escaped the underscore rule and were reported as
   events. The rules now validate the control prefix against designer
   controls (or the intrinsic `UserForm_*` set when controls are
   unknown), mirroring the event_reentry precedent.
5. **Enable by default.** Rejected pending corpus evidence; the rules
   ship opt-in and the corpus workspace enables them for review.

## Evidence

- Issue #826 (sub-issue of #816) defines the required hazard set and the
  opt-in/configurable policy.
- `internal/analyze/class_interface_hazards.go` implements all five
  rules; `internal/analyze/class_interface_hazards_test.go` covers each
  rule's activation boundary, exclusions, interface-member verification,
  form-control recognition, and default-off behavior.
- `docs/specs/vba-class-interface-diagnostics.md` records the public
  rule contract; `internal/staticanalysis/rules/registry.json` carries
  the shared metadata consumed by config, docs, and surface gating.
- VBE oracle fixtures `vba278-public-underscore-member`,
  `vba279-document-public-enum`, `vba280-write-only-property`,
  `vba281-public-event-handler`,
  `vba281-public-interface-member-collision`, and
  `vba282-predeclared-self-access` record the executed Excel
  16.0/17932/x64/ja-JP observations (five accepted, the interface
  collision rejected as ambiguous-name).
- Corpus review: class/interface hazard findings across nine third-party
  projects are committed as reviewed true-positives in
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
