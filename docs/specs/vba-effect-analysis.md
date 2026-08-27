# VBA Procedure Effect Analysis

This specification defines xlflow's deterministic, protocol-neutral procedure
effect summaries. ADR-0023 records the original rationale and ADR-0044 amends
the fixed-point implementation strategy. The summaries are an internal
analysis contract built from resolved procedure IR and conservative CFG facts.
They add no public CLI option or LSP capability by themselves. `VBA221` and
`VBA244` consume the summaries through their own diagnostic contracts;
`VBA221` owns the `detect_application_state_call_effects` compatibility key,
while `VBA244` owns `detect_procedure_call_cycles` and its additive
`call_cycle` finding context.

## Construction and Ownership

`internal/vba/effects` consumes the project documents already parsed and
normalized by `internal/vba/procedureir`, their project resolution overlay, and
the graphs already built by `internal/vba/cfg`. The analyzer constructs those
inputs and the project effect summary once per analysis revision when the
`Effects` capability is required by the active analysis plan. When no enabled
diagnostic requires `Effects`, no project effect summary is constructed. Effect
analysis must not parse source again or retain tree-sitter nodes, source
buffers, CLI envelopes, or LSP protocol values.

### Capability dependency and construction

`Effects` has one explicit project capability dependency: `Resolution`.
Resolution is planned and built before effects, and the effect builder consumes
the shared resolver and resolved IR rather than constructing a private resolver
or project context. The `ApplicationState` and `EventReentry` capabilities
depend transitively on `Effects`; they reuse the same immutable summary.

The capability planner constructs the summary at most once per analysis
revision. Batch analysis may perform that build eagerly after dependency
planning, before procedure workers start. Full realtime/LSP diagnostics may
reuse a revision-scoped immutable cache; a new document revision or a change
between incomplete and complete project evidence invalidates the prior value.
Individual rules must not call an effect builder as a hidden fallback after the
plan has been established.

Whenever an effect consumer is planned, the summary retains the complete
project-local call closure required by the existing propagation and uncertainty
contracts. Procedure participant filtering may be used by other domains, but
it must not reduce the effect closure or remove ambiguous, unresolved, dynamic,
or incomplete call boundaries. If resolution is incomplete, the existing
conservative uncertainty behavior remains in force.

Procedure identity combines the normalized source identity, qualified
procedure name, procedure kind, and declaration location needed to distinguish
otherwise equal VBA names. Project summaries are searchable by that identity.
Returned procedure summaries and every evidence collection use stable ordering
and do not depend on file input or Go map iteration order.

## Model

`EffectKind` has this initial vocabulary:

- `writes_cells`
- `changes_workbook`
- `opens_workbook`
- `opens_file`
- `closes_workbook`
- `disables_events`
- `restores_events`
- `changes_calculation`
- `recalculates`
- `changes_selection`
- `changes_controls`
- `shows_dialog`
- `launches_process`
- `suppresses_errors`
- `raises_error`
- `changes_application_state`
- `restores_application_state`

The last two internal kinds preserve property-neutral `VBA203` compatibility
for `ScreenUpdating`, `EnableEvents`, `DisplayAlerts`, `Calculation`,
`StatusBar`, `Cursor`, `Interactive`, `AskToUpdateLinks`, `AutomationSecurity`,
and `CutCopyMode`. Evidence still carries the exact `Application.<Property>`
target; the more-specific event and calculation kinds are emitted alongside
them where applicable.

`Evidence` identifies the effect, originating procedure identity, canonical
`ast.Range`, statement or call ID, and an optional normalized target such as an
`Application` property. The origin does not change when evidence is propagated
through callers.

Each `ProcedureSummary` separates `Direct` from `Propagated` evidence and
`DirectUncertainty` from `PropagatedUncertainty`. Counts, when needed
internally, are derived from these deduplicated collections rather than stored
as independently mutable totals. `ProjectSummary` contains the sorted procedure
summaries and supports exact stable-identity lookup.

The fixed-point builder keeps a compact semantic state separately from this
compatibility projection. Semantic membership is bounded by the finite effect,
error-outcome, and uncertainty domains. Direct evidence and deterministic
witness seeds are retained for facts that may need an explanation; equivalent
transitive origins, uncertainty entries, and call-chain prefixes are not
eagerly copied through every caller. The projection still exposes the same
direct/propagated distinction and materializes any requested evidence in stable
procedure, source, and edge order.

`BuildWithStats(documents)` returns the project summary and developer-facing
`BuildStats`; `Build(documents)` remains the compatibility wrapper. The stats
fields are `WorklistEvaluations`, `MaxPropagatedFactsPerProcedure`, and
`TotalPropagatedFacts`. They count compact propagated semantic facts, not the
number of entries created when a consumer requests a full provenance view.

For issue #446, `ProcedureSummary` additionally carries an `ErrorSummary`
derived from reachable CFG outcomes. It records provenance for handler
presence, resume-next use, suppression, explicit rethrow, Boolean success
returns, possible raises, and logging followed by continuation. These outcomes
reuse the confirmed call edges and fixed-point rules below but are not inferred
from the `suppresses_errors` or `raises_error` effect kinds alone. The complete
outcome and diagnostic contract is defined in
[VBA Error Suppression and Propagation Analysis](vba-error-propagation-analysis.md).

## Direct Effect Detection

Direct effects are positive semantic claims and therefore require a
high-confidence normalized IR pattern at a conservatively reachable statement.
A statement proven unreachable by its CFG is excluded. A recovered statement,
comment, string literal, or member access through an untyped alias does not
produce positive evidence.

The initial detectors cover:

- assignment to, or clear/insert/delete operations on, a confidently
  recognized `Range`/`Cells` target: `writes_cells` and `changes_workbook`;
- confidently recognized workbook or worksheet structural mutation and save:
  `changes_workbook`;
- `Workbooks.Open`: `opens_workbook`;
- a conservatively reachable, non-recovered VBA `Open ... For ... As #handle`
  acquisition: `opens_file`. This is effect evidence for interprocedural
  consumers; `VBA219` continues to own the separate local ownership and leak
  analysis for the handle and its cleanup paths.
- a confidently recognized `Workbook.Close`: `closes_workbook`;
- every recognized assignment to the ten `Application` properties above:
  `changes_application_state`; known property-reset values and resolved
  saved-value variables additionally carry `restores_application_state` for
  the established Push/Pop helper convention. A state-setting assignment,
  including `Application.EnableEvents = False`, never doubles as restoration
  evidence;
- disabling and restore-candidate assignments to `Application.EnableEvents`:
  `disables_events` or `restores_events`, plus property-neutral evidence; and
- assignments to `Application.Calculation`: `changes_calculation`, plus the
  property-neutral evidence;
- `MsgBox` and `InputBox`: `shows_dialog`;
- `Shell`, and explicit `WScript.Shell.Run` or `WScript.Shell.Exec`:
  `launches_process`;
- `On Error Resume Next`: `suppresses_errors`; and
- `Err.Raise` and the VBA `Error` statement: `raises_error`.

Workbook open and close are lifecycle effects and are not automatically
reclassified as generic `changes_workbook`. Cell writes intentionally produce
both the narrow and aggregate mutation effects.

## Calls and Uncertainty

A call creates a project propagation edge only when all of these conditions
hold:

- its statement is conservatively reachable;
- resolution status is `matched`;
- resolution contains exactly one candidate; and
- that candidate identifies one procedure in the current project summary.

Receiver-less cross-module resolution considers standard-module procedures
only. Class, document, and UserForm procedures require an explicit receiver
outside their own module; sidecar UserForm sources retain their configured
`form` module kind for this decision.

Other reachable calls retain call-site provenance as direct uncertainty:

- `ambiguous` for multiple visible candidates;
- `unresolved` when no known candidate exists;
- `external` for known external dispatch; and
- `dynamic` for `member_call` or `dynamic` resolution that cannot be bound
  conservatively;
- `incomplete` when resolution cannot produce a complete call fact; and
- `non_callable` when a visible candidate is not callable in the current
  context.

`builtin_like` is not project-call uncertainty. It contributes an effect only
when a direct detector above recognizes the built-in; otherwise it is ignored
by this layer. Unreachable calls and recovered call facts do not create edges
or affirmative effects.

## Fixed-Point Propagation

Propagation is finite set-union over the confirmed project-local call edges.
The fixed-point state uses bounded semantic membership and indexed,
deduplicated adjacency. The worklist re-evaluates a caller only when a new
semantic fact, uncertainty member, or required witness seed is observed. The
worklist and every adjacency/evidence collection use stable identity and
source-order keys.

The compatibility projection preserves the existing rule: a caller receives
the callee's direct and propagated effect evidence as propagated evidence, and
likewise receives its uncertainty as propagated uncertainty. That projection
is materialized from compact state and witness seeds when requested; it is not
the state used to decide fixed-point convergence.

Evidence is keyed by its origin procedure and source occurrence, including the
effect or uncertainty kind and relevant target. Consequently, a shared leaf in
a diamond graph appears once in the root summary, and self or mutual recursion
reaches a fixed point without count growth. A procedure's own direct evidence
is never added to that same procedure's propagated collection through a cycle.
Direct and propagated collections remain separate even when they contain the
same effect kind from different origins.

Error outcomes follow the same bounded-state boundary. The finite outcome
membership (`suppresses_errors`, `rethrows_errors`, `returns_success_flag`,
`may_raise`, and `logs_and_continues`) converges independently from detailed
`ErrorEvidence`. Error witnesses retain the handler, cleanup, return, or call
boundary needed by `VBA237`. A shared materialization cache replays the legacy
global procedure worklist once for error witnesses, including cross-origin
requeues, so representative chains preserve the pre-optimization first-arrival
ordering and do not depend on map iteration timing.

## `VBA203` Integration

Issue #428 uses effect summaries only to validate the existing paired
Push/Pop-helper exemption. The assignment that changes `Application` state
remains the root diagnostic site. For a `PushX` procedure, pair lookup prefers
the corresponding `PopX` or `RestoreX` in the same module; otherwise it accepts
only one project-visible candidate. The pair must contain a matching restore
effect either directly or through confirmed propagated calls.

An ambiguous pair, a missing pair, an uncertain call chain, or a summary that
does not contain the matching restore is not treated as safe. The analyzer
continues to emit at most the existing root-assignment finding; it does not emit
caller-level findings under issue #428.

Issue #431 also consumes the CFG directly for procedure-local all-path
restoration. It tracks saved property-value snapshots and their direct variable
copies, then reports a changed assignment when a normal, exceptional,
termination, or unknown exit can retain that change. Cleanup labels and error
handlers therefore participate naturally. Effect summaries remain the narrow
interprocedural Push/Pop exemption only; they do not prove a caller's paths.
At a control-flow join, a clean path is neutral for restore coverage belonging
to an origin that exists only on a dirty path. A recognized saved-value restore
is the cleanup boundary even on the restore statement's exceptional edge;
failures inside that restoration assignment are outside VBA203's lifecycle
model.

The integration preserves the `VBA203` diagnostic ID, severity, ordering,
inline suppression, configuration, and `analyze --json` field shape. Its
message and reason name the affected property and an available uncovered exit
witness. A restore-like direct effect may be included as explanation-only
evidence when a procedure contains one, but it never clears the CFG leak origin
or proves that the saved value is the value being restored on every exit.

## Boundaries

An effect summary establishes possible reachable behavior, not an all-path
guarantee. Issue #431 provides procedure-local all-path restoration through the
CFG. `VBA221` consumes the direct callee summary only when its direct
Application-state evidence matches a `VBA203` leak origin. It reports that
immediate call boundary once per property, names the originating procedure, and
preserves any callee uncertainty. If the direct callee also has a restore-like
effect for that property, the finding may identify it as a candidate while
stating that it is not an all-exit proof. It never repeats the same leak at
transitive ancestors or treats a later restore helper as an all-path proof.

`VBA220` consumes these summaries for supported event handlers. It classifies
cell writes, explicit recalculation, selection changes, workbook lifecycle and
structure changes, and recognized UserForm-control changes as same-event or
broader-chain risks. Only uniquely resolved project-local calls prove a
triggering effect; other calls remain explicit uncertainty. `EnableEvents`
suppresses only Excel events when its False assignment dominates the effect and
the existing CFG analysis proves restoration on every exit; it does not
suppress UserForm control events. For one reachable statement or resolved call
boundary, VBA220 emits at most one finding: a confirmed same-event risk takes
precedence over a broader-chain risk, and either takes precedence over call
uncertainty. The underlying effect summary retains every evidence item for other
consumers.

The CFG retains its existing fallback that every executable statement is a
potential fault site. Effect summaries are not embedded into CFG construction
and do not narrow exceptional edges under issue #428. `VBA237` is the first LSP
consumer of a project effect summary; its complete-workspace cache, Full-only
publication, dependency generations, and old-plus-new reverse-call-graph
invalidation are specified in
[VBA Error Suppression and Propagation Analysis](vba-error-propagation-analysis.md).
