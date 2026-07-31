# VBA Procedure Effect Analysis

This specification defines xlflow's deterministic, protocol-neutral procedure
effect summaries. ADR-0023 records the rationale. The summaries are an internal
analysis contract built from resolved procedure IR and conservative CFG facts;
they do not add a public CLI option, configuration key, JSON field, diagnostic
ID, or LSP capability.

## Construction and Ownership

`internal/vba/effects` consumes the project documents already parsed and
normalized by `internal/vba/procedureir`, their project resolution overlay, and
the graphs already built by `internal/vba/cfg`. The analyzer constructs those
inputs and the project effect summary once per run. Effect analysis must not
parse source again or retain tree-sitter nodes, source buffers, CLI envelopes,
or LSP protocol values.

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
- `closes_workbook`
- `disables_events`
- `restores_events`
- `changes_calculation`
- `shows_dialog`
- `launches_process`
- `suppresses_errors`
- `raises_error`
- `changes_application_state`
- `restores_application_state`

The last two internal kinds preserve property-neutral `VBA203` compatibility
for `EnableEvents`, `DisplayAlerts`, `ScreenUpdating`, and `Calculation`.
Evidence still carries the exact `Application.<Property>` target; the more
specific event and calculation kinds are emitted alongside them where
applicable.

`Evidence` identifies the effect, originating procedure identity, canonical
`ast.Range`, statement or call ID, and an optional normalized target such as an
`Application` property. The origin does not change when evidence is propagated
through callers.

Each `ProcedureSummary` separates `Direct` from `Propagated` evidence and
`DirectUncertainty` from `PropagatedUncertainty`. Counts, when needed
internally, are derived from these deduplicated collections rather than stored
as independently mutable totals. `ProjectSummary` contains the sorted procedure
summaries and supports exact stable-identity lookup.

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
- a confidently recognized `Workbook.Close`: `closes_workbook`;
- disabling and restoring `Application.EnableEvents`: `disables_events` or
  `restores_events`, plus the matching property-neutral state effect;
- changes to `Application.Calculation`: `changes_calculation`, plus a
  property-neutral change or restore classified with the existing `VBA203`
  assignment rules;
- changes and restores of `Application.DisplayAlerts` and
  `Application.ScreenUpdating`: property-neutral state evidence;
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
- `dynamic` for `member_call` resolution that cannot be bound conservatively.

`builtin_like` is not project-call uncertainty. It contributes an effect only
when a direct detector above recognizes the built-in; otherwise it is ignored
by this layer. Unreachable calls and recovered call facts do not create edges
or affirmative effects.

## Fixed-Point Propagation

Propagation is finite set-union over the confirmed project-local call edges.
The worklist and every adjacency/evidence collection use stable identity and
source-order keys. A caller receives the callee's direct and propagated effect
evidence as propagated evidence, and likewise receives its uncertainty as
propagated uncertainty.

Evidence is keyed by its origin procedure and source occurrence, including the
effect or uncertainty kind and relevant target. Consequently, a shared leaf in
a diamond graph appears once in the root summary, and self or mutual recursion
reaches a fixed point without count growth. A procedure's own direct evidence
is never added to that same procedure's propagated collection through a cycle.
Direct and propagated collections remain separate even when they contain the
same effect kind from different origins.

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

The integration preserves the existing `VBA203` diagnostic position, message,
reason, suggestion, severity, ordering, inline suppression, and
`[analyze].disabled_rules` behavior. The `analyze --json` shape is unchanged.

## Boundaries

An effect summary establishes possible reachable behavior, not an all-path
guarantee. Verifying restoration on every exit is issue #431, event re-entry is
issue #438, and caller-level state propagation/reporting is issue #439.

The CFG retains its existing fallback that every executable statement is a
potential fault site. Effect summaries are not embedded into CFG construction
and do not narrow exceptional edges under issue #428. LSP workspace-generation
cache and invalidation behavior will be specified when an LSP effect consumer
is introduced.
