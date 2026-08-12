# ADR-0039: Procedure and Module Hotspot Ranking

## Status

Accepted

## Context

Issue #461 asks xlflow to identify procedures and modules with unusually high
responsibility or dependency centrality. A single complexity threshold cannot
represent call centrality, Excel effects, mutable state, cycles, public API
surface, error handling, or resource ownership. Conversely, presenting a
subjective composite as a definite analyzer finding would make maintainability
triage look like a correctness diagnosis.

ADR-0036 established `xlflow metrics` as a source-only, deterministic surface
for procedure measurements. It deliberately left module/project aggregates,
recommended thresholds, and history outside v1. The hotspot projection needs
those aggregate and dependency signals while retaining the independent metrics
compatibility boundary.

## Decision

Extend `xlflow metrics` with a versioned `metrics.hotspots` projection. The
projection ranks procedure and module cohorts independently and includes both
raw scalar signals and the normalized values that contribute to each score.
The source scan remains read-only and source-only: no Excel/COM operations,
diagnostic enablement, LSP projection, or analyzer rule execution is required.

The v1 score model is named `percentile_equal_weight_v1`. For each signal in a
cohort, values are converted to average-rank percentiles. Constant and
single-entity signals are inactive; the score is the equal-weight mean of the
remaining signals, rounded to two decimals. Scores and arrays are sorted
deterministically, with stable identity tie-breakers. A new score interpretation
requires a new model identifier rather than silently changing historical scores.

The raw signal vocabulary is additive and conservative: structural complexity,
module complexity maximum, resolved call fan-in/out, affected modules, confirmed Excel effects, mutable
state reads/writes/mutations, cycle participation, external dependencies,
error-handling structures, resource ownership, and module public-procedure
count. Missing, unresolved, ambiguous, external, or dynamic facts contribute no
invented project-local dependency. Ambiguous, unresolved, and dynamic call
counts remain available in an `uncertainty` evidence map but are excluded from
the score.

Selection is opt-in under `[metrics.hotspots]`, independently for procedures and
modules. `procedure_top_n` and `module_top_n` select ranked entities; positive
`*_score_threshold` values select scores greater than or equal to the configured
percentage. The two modes are unioned and the JSON `selected_by` field records
the reason. Selected entities emit metrics-owned `MX002`: top-N-only output is
informational, while threshold-selected output is warning-level and exits `1`.
Threshold findings retain the complete metrics/hotspot payload. `MX002` is not a
lint/analyze rule and cannot be inline-suppressed.

## Consequences

- Reports, CI jobs, and editor integrations can rank architectural review leads
  without parsing diagnostic prose or enabling runtime-risk rules.
- Raw and normalized evidence makes every score auditable and allows projects to
  calibrate selectors against real distributions before adopting a gate.
- Percentile normalization avoids coupling v1 to arbitrary cross-project units,
  but scores are relative to the current procedure/module cohort and are not
  comparable across snapshots without preserving the source population.
- Equal weighting is intentionally simple and transparent; domain-specific
  weighting, historical baselines, trend detection, and LSP projections remain
  future work and require a new decision or model version.
- Conservative dependency/effect resolution avoids false certainty but can
  under-rank code whose dynamic behavior cannot be resolved statically.
- Warning-level threshold selection can fail CI only when a project opts in;
  raw rankings and informational top-N output remain non-blocking.

## Alternatives Considered

1. **Add a complexity-only `MX002` threshold.** Rejected because one metric
   cannot represent centrality, effects, state, cycles, or ownership and would
   recreate the arbitrary-threshold problem from Issue #461.
2. **Register hotspots as an `analyze` rule.** Rejected because hotspot scores
   are measurements and prioritization aids, not definite runtime-risk
   diagnostics; coupling them to analyzer enablement would make output and exit
   behavior unstable.
3. **Use hand-tuned weights or a single global threshold.** Rejected for v1:
   weights require reviewed project evidence and a policy owner. Percentile
   equal weighting is inspectable and can evolve through an explicit model ID.
4. **Count every unresolved or external call as a dependency.** Rejected
   because uncertainty is not evidence of a project-local relationship and
   would inflate scores in dynamic or incomplete projects.
5. **Create a separate `hotspots` command.** Rejected because the projection
   consumes the same source snapshot and procedure/call facts as `metrics`;
   embedding it keeps one deterministic scan and one versioned envelope while
   leaving room for a future dedicated report command.

## Evidence

- Issue #461 acceptance criteria for composite scores, raw signals, top-N and
  threshold selection, deterministic ranking, and informational/warning output.
- `docs/specs/vba-procedure-complexity-metrics.md` and ADR-0036 for the
  source-only metrics boundary and schema/versioning rules.
- `docs/specs/vba-call-graph-reachability.md` and ADR-0035 for conservative
  project-local call resolution and deterministic ordering.
- `docs/specs/vba-module-state-analysis.md` for mutable-state evidence and
  uncertainty boundaries.
- `internal/vba/hotspots` for percentile ranking, equal-weight scoring, and
  stable tie-break implementation, with deterministic ranking tests in
  `internal/vba/hotspots/hotspots_test.go`.

## Supersedes

None.

## Superseded by

None.

## Related

- Issue #461
- ADR-0035, ADR-0036
- `docs/specs/vba-procedure-and-module-hotspots.md`
