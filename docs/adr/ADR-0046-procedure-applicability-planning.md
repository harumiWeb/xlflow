# ADR-0046: Procedure Feature Sets and Applicability-Based Analysis Planning

## Status

Accepted

## Context

The shared procedure IR, immutable analyzer facts, control-flow graph, effect
summaries, and bounded procedure execution now give `analyze` reusable inputs
for each procedure. The procedure-local phase still enters every enabled
semantic rule family, however. A rule that can prove immediately that its
domain is irrelevant still pays scheduling, setup, lookup, and allocation costs
for every procedure.

Issue #695 addresses that cost without changing diagnostic meaning. The
optimization must also work for realtime/LSP analysis, preserve the existing
batch and project-wide effect contracts, and retain the conservative behavior
of recovered or unresolved VBA.

Issue #696 applies the same planning boundary to overlapping array diagnostics.
Array applicability must decide whether the array domain is needed, while the
domain itself must materialize one canonical procedure result that can feed
several rule projections. Applicability planning must not become a second
array parser or a rule-specific cache.

## Decision

Derive one immutable, procedure-local applicability summary from the owned IR
and `procedureAnalysisFacts` for each analysis revision. The summary uses a
compact representation equivalent to two bit sets: one for features proven
present and one for features whose absence is unknown. A feature in neither
set is proven absent. The concrete bit assignment is internal and may evolve
with the semantic domains; it is not a public CLI or JSON contract.

Feature construction is part of the existing facts construction path. It uses
one bounded pass over the owned declarations, statements, expressions, calls,
and accesses. The planner separately combines that local summary with available
project/module facts. Neither step reparses source or performs a separate
source scan for each feature. Features are added only when
an enabled or planned consumer has a real prerequisite. Current consumers may
classify arrays and `ReDim`, loops, object/member and Dictionary/Collection
access, error handling, runtime candidates, dataflow and command/SQL/HTTP/file
operations, resource ownership, Excel operations, application-state mutation,
event metadata, calls, and ByRef or dynamic-call uncertainty.

The planner maps enabled diagnostic IDs to explicit semantic-domain
requirements, then combines those requirements with the procedure summary,
resolution/CFG completeness, module facts, and propagated effect summaries.
Domain setup and flow-state allocation occur only after this plan is available.
Only a feature/domain proven absent can make a domain skipped. An unknown
feature, incomplete or recovered IR, missing project/module/type facts,
ambiguous or unresolved resolution, dynamic call, or uncertain project effect
keeps the domain planned. A domain with any planned procedure still builds its
complete project-wide index and effect closure; the planner must not construct
a procedure subset that weakens propagation.

The same requirements and planner are used by batch and realtime/LSP entry
points. Compile-equivalent diagnostics, source-line checks, trace-helper
checks, ByRef/file-level Intel checks, and other unconditional compatibility
paths remain outside the gated domains unless a separate decision explicitly
changes them. Planning is an execution optimization only: finding content,
range, multiplicity, ordering, severity, suppression, exit status, and normal
JSON/LSP schemas remain unchanged.

When the array domain is planned, its procedure worker materializes one
immutable `ArrayAnalysisResult` for the current analysis revision and passes
that result to the applicable array projectors. The result reuses the owned
procedure/module facts and is discarded with the procedure analysis; it is not
promoted to a project-wide cache. `array_kernel_runs` and the main
`array_cfg_walks` therefore remain one per applicable procedure even when
`VBA227`, deterministic `VBA249`, `VBA208`, `VBA209`, object-array
`VBA101`/`VBA102`, `VBA241`, and `VBA226` are enabled together. The
`array_projection_runs` counter records applicable projectors, while an
explicit `VBA226` secondary shape pass is the only additional array walk
allowed by this plan. `VBA241` reads shared facts and does not start a new
fixed-point traversal. Rule-specific conservatism remains in the projection
policy, including the source-line lifecycle policy of `VBA227` and the
independent `Range.Value` shape policy of `VBA226`.

The analyzer's opt-in performance recorder emits additive planned/skipped
counters for each gated semantic domain. A counter is recorded once per
procedure/domain decision; unknown applicability is counted as planned. These
counters are stderr performance-log telemetry and do not appear in normal
analysis output.

## Rationale

The IR and fact ownership boundaries already make procedure inputs immutable
and reusable. Recording applicability beside those facts amortizes the cheap
negative proof across every enabled rule family while keeping rule-specific
state and effects out of the shared summary. A separate planner table makes
the prerequisite policy reviewable and prevents each rule entry point from
inventing its own conservative fallback. The present/absent/unknown model
allows the optimization to remove only work that the analyzer can prove cannot
produce a finding.

## Consequences

- Scalar-only procedures can avoid array, dataflow, Excel, and other semantic
  setup when the corresponding requirements are proven absent.
- Large projects expose whether the planner is effective through stable
  planned/skipped counters and existing domain timings.
- Feature construction and planner evaluation add a small per-procedure cost,
  especially for small modules; benchmark coverage must keep that overhead
  visible.
- Conservative unknown handling may leave work enabled for incomplete or
  dynamic VBA. This is intentional and protects findings at the expense of
  some missed optimization opportunities.
- Requirement metadata becomes an internal completeness contract: a new
  gated rule must declare its domain requirements and add planner/correctness
  coverage before it can rely on skipping.
- Array applicability is a domain-level decision rather than one decision per
  array diagnostic. Adding an array projector changes projection work and
  telemetry, not the canonical kernel or main CFG walk count.

## Alternatives Considered

1. **Continue invoking every enabled domain for every procedure** - Rejected
   because repeated setup and allocation cost scales with procedure count even
   when the domain has no possible input.
2. **Reparse or rescan source independently for every domain** - Rejected
   because it duplicates parser/IR work and risks disagreement between rule
   families; the shared owned IR is the source of truth.
3. **Treat an unrecognized or recovered feature as absent** - Rejected because
   unsupported syntax, unresolved calls, and missing project facts can still
   produce a relevant finding; unknown must fail open.
4. **Use one coarse `has-semantic-work` flag** - Rejected because it cannot
   skip independent domains such as arrays versus dataflow and would hide
   which prerequisite justifies a planned run.
5. **Restrict planning to batch analysis or change the public rule registry** -
   Rejected because batch and realtime/LSP must share applicability semantics,
   while the registry's public schema should remain focused on rule metadata.

## Evidence

- Requirement: xlflow issue #695, including the procedure feature, planner,
  fail-open, telemetry, and ROneCOne acceptance criteria.
- Immutable IR and recovery/ownership contract:
  `docs/specs/vba-analysis-ir.md` and
  `docs/adr/ADR-0021-procedure-analysis-ir.md`.
- CFG and project effect boundaries:
  `docs/specs/vba-control-flow-graph.md`,
  `docs/specs/vba-effect-analysis.md`, and
  `docs/adr/ADR-0023-procedure-effect-summaries.md`.
- Existing procedure facts, semantic-domain entry points, and performance
  recorder: `internal/analyze/procedure_facts.go`,
  `internal/analyze/analyzer.go`, and
  `internal/analyze/profiling_test.go`.
- Existing analyzer performance and corpus measurement procedure:
  `docs/specs/cli-contract.md` and
  `docs/specs/static-analysis-corpus.md`.
- ADR-0040 for the canonical array semantic result and its diagnostic
  ownership boundaries.

## Related

- Issue #693 (parent)
- Issue #695
- Issue #696
- ADR-0021, ADR-0022, ADR-0023, ADR-0024, ADR-0043, ADR-0045
- ADR-0040
- `docs/specs/vba-analysis-ir.md`
- `docs/specs/static-analysis-corpus.md`
