# ADR-0023: Deterministic Procedure Effect Summaries

## Status

Accepted

## Context

ADR-0021 provides normalized procedure syntax and conservative call
resolution. ADR-0022 adds procedure-local reachability, but neither layer can
answer whether a procedure performs a relevant side effect itself or only
through another project procedure. Issue #428 requires that information as a
shared foundation for the reliability-analysis program in issue #425.

An interprocedural summary cannot treat every call as a confirmed dependency.
VBA permits ambiguous project names, external dispatch, built-in calls, and
member calls whose receiver type is unknown. Recursion, mutual recursion, and
diamond-shaped call graphs also make traversal order and duplicate evidence
observable unless propagation has an explicit convergence contract.

The first consumer is deliberately narrow. `VBA203` already recognizes paired
Push/Pop or Push/Restore helpers for Excel `Application` state. It needs to
recognize a restore reached through a uniquely resolved helper without changing
its root-cause diagnostic contract. Issue #439 additionally needs one useful
caller context for a direct helper that can leave Application state changed.

## Decision

Add `internal/vba/effects` as a protocol-neutral layer above
`procedureir` and `cfg`. It builds one project summary from already-built,
resolved procedure IR and CFG values; it does not parse source or retain parser,
CLI, JSON-envelope, or LSP values.

Represent direct effects as finite, provenance-bearing evidence. The initial
effect vocabulary covers cell writes, workbook mutation and lifecycle,
`Application.EnableEvents` and `Application.Calculation` state, dialogs,
process launch, and error suppression or raising. Positive evidence is emitted
only for high-confidence normalized patterns that are not recovered and are not
proven unreachable by the CFG.

Preserve the existing four-property `VBA203` compatibility contract with two
additional internal property-neutral kinds for Application-state change and
restore. Their evidence identifies the exact property, while EnableEvents and
Calculation also retain their more specific effect kinds.

Represent ambiguous, unresolved, external, and dynamically bound member calls
as explicit uncertainty with call-site provenance. Built-in-like calls are not
project-call uncertainty; a recognized built-in may still contribute a direct
effect. Only a `matched` call with exactly one project-local candidate and a
reachable call site becomes a propagation edge. Configured module kind is part
of resolution context: receiver-less cross-module calls consider standard
modules only, while class, document, and UserForm procedures require an
explicit receiver outside their own module.

Compute propagated effects and propagated uncertainty by deterministic
set-union to a fixed point. Evidence identity includes its originating
procedure and source occurrence, so recursion and diamond graphs converge and
do not inflate counts. A procedure's own direct evidence remains direct and is
not reintroduced as propagated evidence through a recursive cycle. Project and
procedure summaries are sorted by stable procedure identity and stable
provenance keys, never input or map iteration order.

Use the summary as a validation consumer for the existing
`VBA203` Push/Pop helper exemption. Pair lookup prefers the same module, then
requires exactly one project-visible Pop/Restore candidate. A direct or
propagated matching restore effect permits the existing exemption; missing or
ambiguous pairs remain unsafe. Diagnostic ID, severity, location, message,
ordering, inline suppression, configuration, and CLI/LSP representation remain
unchanged.

`VBA221` reports a batch-only caller context at a uniquely resolved, reachable
call only when the direct callee has direct Application-state evidence that
matches a CFG-proven `VBA203` leak origin. It names that property and origin,
and includes propagated call uncertainty from the callee. It does not report a
transitive callee leak at every ancestor, and it does not use later restoration
helpers to claim that the caller is safe. `VBA203` remains at the direct
assignment that introduced the leak.

Event re-entry analysis and workspace/LSP effect caching remain outside this
decision. Issue #431 still consumes the CFG directly to prove procedure-local
restoration on every exit; summaries do not establish that proof. The CFG keeps
its existing conservative rule that every executable statement may fault;
effect summaries do not alter that fallback.

## Consequences

- Positive: later reliability checks can reuse one deterministic project-wide
  effect model instead of adding rule-specific call traversal.
- Positive: source evidence remains attributable after propagation, while
  unresolved dispatch stays visible rather than becoming false certainty.
- Positive: finite-set propagation terminates for recursive, cyclic, and
  diamond-shaped call graphs and produces stable counts and ordering.
- Positive: `VBA203` can validate an indirect Pop/Restore helper without adding
  a new diagnostic or public data contract.
- Positive: `VBA221` gives the immediate caller actionable context while the
  original assignment remains the single root-cause location.
- Negative: high-confidence detectors intentionally miss effects hidden behind
  aliases, late binding, type-library dispatch, or recovered syntax.
- Negative: retaining provenance uses more memory than one boolean per effect.
- Limitation: a summary reports that an effect is reachable. The CFG-based
  procedure-local #431 analysis separately proves local exit restoration, but
  neither layer makes an arbitrary caller or dynamic helper path safe.

## Alternatives Considered

1. **Propagate booleans or counts only** - Rejected because call chains could
   not be explained, recursion could inflate counts, and duplicate evidence in
   diamond graphs would become traversal-order dependent.
2. **Treat every syntactic call as an edge** - Rejected because ambiguous,
   external, and dynamically dispatched calls are not proof of one local
   callee and would create false positive effect claims.
3. **Embed effects in the procedure IR or CFG** - Rejected because syntax,
   control flow, and interprocedural semantic aggregation have different
   ownership and refresh boundaries.
4. **Implement only a `VBA203` helper search** - Rejected because it would
   duplicate resolver and graph traversal policy and would not provide the
   reusable issue #425 foundation.
5. **Report every transitive caller** - Rejected because each ancestor would
   repeat the same root cause without adding a nearer actionable call boundary.

## Evidence

- Requirements: xlflow issues #425, #428, and #439.
- Syntax and resolution foundation:
  `internal/vba/procedureir`, `docs/specs/vba-analysis-ir.md`.
- Reachability foundation:
  `internal/vba/cfg`, `docs/specs/vba-control-flow-graph.md`.
- Existing conservative call resolution:
  `internal/vba/calls`, `internal/vba/callgraph`.
- Validation consumer and compatibility tests:
  `internal/analyze/analyzer.go`, `internal/analyze/analyzer_test.go`.
- Detailed effect contract: `docs/specs/vba-effect-analysis.md`.

## Related

- `docs/adr/ADR-0013-analyze-runtime-risk-ownership.md`
- `docs/adr/ADR-0014-reusable-vba-lsp-server.md`
- `docs/adr/ADR-0021-procedure-analysis-ir.md`
- `docs/adr/ADR-0022-conservative-vba-control-flow-graph.md`
- xlflow issues #425, #428, #431, #438, and #439
