# VBA Procedure and Module Hotspots

This specification defines the architectural-hotspot projection embedded in
`xlflow metrics`. It extends the procedure-complexity contract without turning
hotspots into an analyzer rule or a claim that a procedure is defective.

## Scope and command

```text
xlflow [--json] metrics
```

Hotspots are collected during the same source-only scan as procedure metrics.
The command does not open Excel, use the bridge, execute VBA, run `lint`,
`analyze`, `check`, LSP diagnostics, or preflight. Source discovery,
UserForm-sidecar handling, and `[metrics].exclude` follow
[VBA Procedure Complexity Metrics](vba-procedure-complexity-metrics.md).

The report ranks procedures and modules in separate cohorts. It is emitted in
the `metrics.hotspots` object even when no selector is configured, so reports
and editors can consume the raw evidence without enabling a subjective gate.

## Report schema

The additive `hotspots` projection has its own schema version and score-model
identifier:

```json
{
  "metrics": {
    "schema_version": 1,
    "procedures": [],
    "hotspots": {
      "schema_version": 1,
      "score_model": "percentile_equal_weight_v1",
      "procedures": [
        {
          "id": "src/modules/Main.bas|Main|Run|sub",
          "kind": "procedure",
          "file": "src/modules/Main.bas",
          "module": "Main",
          "module_kind": "standard",
          "name": "Run",
          "procedure_kind": "sub",
          "line": 1,
          "rank": 1,
          "score": 87.5,
          "score_model": "percentile_equal_weight_v1",
          "active_signal_count": 4,
          "raw_signals": {
            "complexity": 8,
            "call_fan_in": 3,
            "call_fan_out": 4,
            "affected_module_count": 2
          },
          "normalized_signals": {
            "complexity": 100,
            "call_fan_in": 75,
            "call_fan_out": 75,
            "affected_module_count": 100
          },
          "uncertainty": {
            "ambiguous_call_count": 0,
            "unresolved_call_count": 1,
            "dynamic_call_count": 0
          },
          "selected_by": { "top_n": true }
        }
      ],
      "modules": []
    }
  }
}
```

`id`, paths, module and procedure identity fields are project-relative and use
`/` separators. `rank` is one-based, with the highest score first. Procedure
entities include `procedure_kind` and may include a one-based declaration
`line`; module entities omit `procedure_kind`. `raw_signals`
contains non-negative integer evidence and `normalized_signals` contains the
percentile values that contributed to the score. `active_signal_count` is the
number of non-constant signals in that cohort. `selected_by` is present only
for entities selected by `[metrics.hotspots]` and records the union reason:
`top_n` and/or `threshold`.

`uncertainty` exposes ambiguous, unresolved, and dynamic call counts. These
values are retained as audit evidence and are deliberately excluded from the
composite score; uncertain relationships are never treated as confirmed
project-local dependencies.

`hotspots.schema_version` is `1`. Additive fields and signals are compatible;
removing a field or changing a signal's meaning or score calculation requires a
new schema version or score-model identifier. Consumers must not infer a
meaning for an unknown version or model.

## Raw signals

Signals are intentionally exposed so a score is auditable. A missing or
uncertain fact contributes `0`; unresolved, ambiguous, external, and dynamic
relationships do not become invented project-local dependencies.

| Signal                      | Evidence represented                                                                                                     |
| --------------------------- | ------------------------------------------------------------------------------------------------------------------------ |
| `complexity`                | Structural procedure complexity from the metrics IR/CFG.                                                                 |
| `complexity_max`            | Maximum procedure complexity in a module (module cohort only).                                                           |
| `call_fan_in`               | Unique resolved project-local callers.                                                                                   |
| `call_fan_out`              | Unique resolved project-local callees.                                                                                   |
| `affected_module_count`     | Distinct project modules reached through resolved project-local call dependencies, including the entity's origin module. |
| `excel_effect_count`        | Confirmed Excel object-model effects.                                                                                    |
| `mutable_state_reads`       | Reads of indexed mutable module state.                                                                                   |
| `mutable_state_writes`      | Writes to indexed mutable module state.                                                                                  |
| `mutable_state_mutations`   | Confirmed mutable-state mutations.                                                                                       |
| `cycle_count`               | Distinct detected project-local call cycles containing the entity.                                                       |
| `external_dependency_count` | External or declared dependency references retained as scalar evidence.                                                  |
| `error_handling_count`      | Error-handler and failure-handling structures.                                                                           |
| `resource_ownership_count`  | Resource-acquisition/ownership boundaries.                                                                               |
| `public_procedure_count`    | Public procedures in a module (module cohort).                                                                           |

Signal values are counts, not diagnostic severities. A high score is a
maintainability-review lead, not proof of a bug, security issue, or required
refactor. The collector remains conservative when evidence is incomplete. To
keep source-only metrics bounded on dense graphs, elementary-cycle enumeration
uses a deterministic fixed work budget; if it is exhausted, `cycle_count` is a
conservative lower bound for that report.

## Score and ordering

Procedure and module cohorts are normalized independently. For a cohort of
`N >= 2` entities, each signal is sorted ascending and a value tied from the
1-based ordinal `r_first` through `r_last` receives
`average_rank = (r_first + r_last) / 2`, then
`percentile = 100 * (average_rank - 1) / (N - 1)`. Thus the lowest and highest
values are exactly `0` and `100`, and ties share their average-rank percentile.
A signal whose values are constant (or whose cohort has one entity) is
inactive, contributes percentile `0`, and is omitted from the score
denominator. The score is the equal-weight arithmetic mean of active
normalized signals, rounded to two decimal places; entities with no active
signals score `0`. Score-threshold selection compares the emitted two-decimal
score (`score >= threshold`), not an unrounded intermediate value.

Entities sort by descending score. Equal scores use normalized file,
declaration byte, module, name, procedure kind, and stable `id` as tie-breakers.
The resulting arrays and rank values are independent of source enumeration,
map iteration, path separator, or line-ending order.

## Selection and `MX002`

Selectors are opt-in and configured independently for procedures and modules:

```toml
[metrics.hotspots]
procedure_top_n = 0
module_top_n = 0
procedure_score_threshold = 0
module_score_threshold = 0
```

`*_top_n` must be a non-negative integer; `0` disables it. A positive value
selects that many highest-ranked entities (or all entities when fewer exist).
`*_score_threshold` is a percentage in the inclusive range `1..100`; `0`
disables it. An entity is selected when `score >= threshold`. Top-N and
threshold selections are unioned, and `selected_by` records every reason.

Each selected entity produces one `MX002` diagnostic with the entity identity,
score, score model, raw signals, and selection reason. Top-N-only selections
are informational and do not fail the command. A selection caused by a score
threshold is warning-level (including an entity selected by both modes) and
causes exit code `1`; the complete metrics and hotspot arrays remain in the
failure envelope. `MX002` is owned by `metrics`, is not in the analyzer rule
registry, and is not inline-suppressible.

Diagnostics follow ranked procedure order, then ranked module order. Their
structured fields mirror the selected entity (`kind`, `file`, `line`, `module`,
`procedure`, `rank`, `score`, `score_model`, signal maps, and `selected_by`).

Invalid selector values, unknown keys, malformed tables, negative top-N values,
non-finite selector percentages, or thresholds outside `0..100` are
configuration errors (exit `2`). A non-finite computed score is an internal
failure and must not publish partial metrics or hotspot rankings. Source
read/parse failures fail closed without publishing partial metrics or hotspot
rankings.

## Determinism and consumers

The hotspot report is derived from the same source snapshot as procedure
metrics. Consumers should persist both `schema_version` and `score_model`, keep
the raw and normalized signal maps, and treat unsupported versions/models as
unavailable rather than guessing. Hotspots are a prioritization aid; use
`xlflow analyze` for definite runtime-risk diagnostics and use `MX001` only for
explicit procedure metric limits.

## Related

- `docs/adr/ADR-0039-procedure-and-module-hotspots.md`
- `docs/specs/vba-procedure-complexity-metrics.md`
- `docs/specs/vba-call-graph-reachability.md`
- `docs/specs/vba-module-state-analysis.md`
- `vitepress/reference/json-output.md`
