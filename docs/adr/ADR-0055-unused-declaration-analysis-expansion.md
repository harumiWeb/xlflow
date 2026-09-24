# ADR-0055: Unused Declaration Analysis Expansion

## Status

Accepted

## Context

Issue #821 (sub-issue of #816) asks xlflow to cover Rubberduck-style unused
declaration analysis for procedure parameters, constants, user-defined type
members, broader unreachable procedures, and definite-assignment checks for
local variables. The existing analyzer already computes `VBA256` dead stores
from backward liveness over the procedure CFG and reports `VB021` for
unreachable private procedures, but parameter/constant/UDT-member dead code
and use-before-assignment have no coverage.

The open design questions are which visibility boundary makes an "unused"
claim sound, how much semantic resolution each declaration class needs before
it can fail open, and whether unassigned-variable detection belongs in the
existing definite-assignment machinery or a separate pass.

## Decision

Add five opt-in analyzer rules plus one extension to an existing lint rule.

| Rule     | Contract                                                                                                                                                     | Default          |
| -------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------ | ---------------- |
| `VBA265` | A parameter of a `Private` procedure that is never referenced in the procedure body is reported as a warning.                                                | Disabled; opt-in |
| `VBA266` | A module-level `Private Const` never referenced by any expression, declaration initializer, or conditional-compilation directive in the module is a warning. | Disabled; opt-in |
| `VBA267` | A member of a `Private Type` that no resolvable member expression in the module touches is an information finding.                                           | Disabled; opt-in |
| `VBA268` | A procedure-local scalar variable that is read but never assigned is a warning.                                                                              | Disabled; opt-in |
| `VBA269` | A procedure-local scalar variable read on a reachable path without a guaranteed assignment is a warning.                                                     | Disabled; opt-in |
| `VB021`  | Extended to `Friend` procedures and to `Public` procedures inside host-hidden (`Option Private Module`, non-`VB_Exposed`) modules.                           | Disabled; opt-in |

Visibility is the soundness boundary throughout. `Private` declarations are
file-local, so `VBA266`/`VBA267` need no project view and run in realtime.
`VBA265` restricts to `Private` procedures because `Public`/`Friend`
signatures are fixed by unenumerable callers; event handlers, `Implements`
members, `Declare` statements, and procedures named by a statically
discoverable string literal (`Application.Run` and friends) are excluded.
`VB021` treats `Friend` and host-hidden `Public` procedures as
project-internal: they have no host-facing surface, so unreachability is
provable, while `VB_Exposed` modules and document/form surfaces stay roots.

`VBA268`/`VBA269` share one candidate pass over `ProcedureIR` accesses scoped
to plain scalar locals. A `ByRef` (or unresolved) call argument counts as a
potential definition: it suppresses "never assigned" and joins the CFG's
definite-assignment fixpoint, but the argument position itself is not a
proven read, because initializer-style calls only write. Unknown flow
sources, uncertain edges, recovered procedures, conditional-compilation
branches, and `LSet`/`RSet`/`Mid$`/`Input #` shapes fail open per procedure
or per variable.

`VBA267` resolves member usage through qualified and `With`-implicit member
expressions, nested UDT member types, and same-module function return types.
An unresolved receiver marks the member name used on every private UDT that
declares it, and a type escaping through `Variant`/object/`ByRef`/I/O
boundaries is exempted entirely.

The corpus workspace enables all five flags so real-world evidence
accumulates without changing production defaults.

## Consequences

- Dead parameters, constants, UDT members, and uninitialized reads surface
  as opt-in warnings without changing any default-enabled diagnostic.
- `VBA265`/`VBA266` use lexical whole-word scans that treat comments and
  string literals as uses; that only suppresses findings and cannot create
  false positives, at the cost of occasionally missing a name that appears
  only in prose.
- `VBA268`/`VBA269` cover only scalar locals; objects, Variants, arrays,
  statics, and UDT variables stay silent because their initialization
  semantics are not provable at whole-variable granularity.
- `VB021` now reports `Friend` and host-hidden `Public` procedures under the
  same diagnostic ID and configuration key, so enabling it in an existing
  project can surface additional true positives.
- `callgraph.ReachabilityRequest` gained a `Reportable` node set; callers
  that leave it nil keep the previous private-only `Unreachable` contract.

## Alternatives Considered

1. **A new diagnostic ID for friend/host-hidden procedures.** Rejected
   because the finding is the same defect with the same remediation;
   splitting IDs would only complicate suppression and review.
2. **Semantic use detection for `VBA265`/`VBA266` via `VariableAccess`
   scopes.** Rejected for the declaration-name scan because procedure IR
   accesses do not cover module-level initializers, attributes, or
   conditional-compilation directives; a lexical whole-word scan covers all
   of them uniformly and errs toward suppression.
3. **Whole-project UDT member resolution.** Rejected because `Private Type`
   is file-local; widening the scope adds a project dependency without
   improving precision.
4. **Count `ByRef` arguments as reads for `VBA269`.** Rejected because
   initializer-style calls are common and only write; the argument position
   stays silent while earlier reads are still reported.
5. **Enable the new rules by default.** Rejected pending corpus evidence;
   the flags ship opt-in and the corpus workspace enables them for review.

## Evidence

- Issue #821 (sub-issue of #816) defines the required declaration classes
  and the conservative external-entry requirements.
- `internal/analyze/unused_parameter.go`, `unused_private_const.go`,
  `unused_udt_member.go`, and `variable_assignment.go` implement the rules;
  `internal/analyze/unused_declaration_test.go` covers each category.
- `internal/vba/reachability/reachability.go` computes the reportable set and
  host-hidden module facts; `internal/vba/callgraph/callgraph.go` accepts
  the `Reportable` overlay for `Unreachable` classification.
- `docs/specs/vba-unused-declaration-diagnostics.md` records the public rule
  contract; `internal/staticanalysis/rules/registry.json` carries the shared
  metadata consumed by config, docs, and surface gating.
- `docs/adr/ADR-0024-shared-static-analysis-rule-registry.md` and
  `docs/adr/ADR-0054-discarded-function-return-diagnostics.md` establish the
  registry, fail-open resolution, and opt-in policy patterns this change
  follows.

## Related

- Issue #816 (parent)
- Issue #821
- ADR-0024, ADR-0046, ADR-0054
- `docs/specs/vba-unused-declaration-diagnostics.md`
- `docs/specs/vba-call-graph-reachability.md`
- `docs/specs/vba-dead-store-diagnostics.md`
