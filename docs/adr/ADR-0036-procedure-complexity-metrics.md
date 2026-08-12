# ADR-0036: Procedure Complexity Metrics

## Status

Accepted

## Context

Issue #460 needs maintainability measurements for VBA procedures. The existing
`lint` and `analyze` surfaces are diagnostic surfaces: their findings are
rule-owned, configurable through the analyzer registry, and may change as
runtime-risk analysis evolves. Putting subjective complexity thresholds into
those surfaces would make a measurement useful only when a particular finding
is enabled and would couple consumers to diagnostic rule ownership. Existing
`analysis_metrics.module_state` is likewise an analyzer-owned projection and
must retain its compatibility contract.

Consumers instead need a deterministic, source-only procedure inventory that
can be stored, compared, and used by future reports without opening Excel or
parsing diagnostic prose. The calculation must also have one interpretation of
single-line, colon-separated, continued, property, label, and generated
UserForm syntax.

## Decision

Add a source-only `xlflow metrics` command. It reuses the parsed procedure IR,
control-flow graph, and uniquely resolved project-local call edges, but has its
own collector and threshold evaluator. It does not run `lint`, `analyze`,
`check`, LSP diagnostics, preflight rules, or Excel/COM operations. The command
scans configured VBA source roots and `tests`; generated build state and
sidecar-generated `.frm` code are not scanned. In UserForm `sidecar` mode the
sidecar code file is the procedure source. A separate `[metrics].exclude`
glob list is applied after normal source discovery and is not inherited from
`[build].exclude` or `[analyze].disabled_rules`.

The command's machine-readable result is a versioned `metrics` object with
`schema_version: 1` and a `procedures` array. Each procedure carries its
project-relative slash-separated file, module and module kind, procedure name
and kind, and a `declaration_range` object with one-based `startLine` and
`endLine` (plus the parser's column/byte offsets), and these twelve integer
metrics:

- `cyclomatic_complexity`
- `max_nesting_depth`
- `statement_count`
- `source_line_count`
- `branch_count`
- `loop_count`
- `goto_count`
- `exit_point_count`
- `parameter_count`
- `byref_parameter_count`
- `local_variable_count`
- `call_fan_out`

The counting rules are permanent and are defined in
`docs/specs/vba-procedure-complexity-metrics.md`. In particular, cyclomatic
complexity is `1 + branch_count + loop_count`; unknown CFG edges, Boolean
operators, and `GoTo` do not add branches. `call_fan_out` contains only unique,
resolved project-local callees. Source line and statement counts deliberately
have different units: physical lines include comments, blanks, and
continuations between the procedure header and matching `End`, while
statements use normalized source-backed IR (colon-separated statements are
separate and a continuation is one statement).

Thresholds are opt-in and independent of diagnostics:

```toml
[metrics]
exclude = []

[metrics.thresholds]
cyclomatic_complexity = 0
max_nesting_depth = 0
statement_count = 0
source_line_count = 0
branch_count = 0
loop_count = 0
goto_count = 0
exit_point_count = 0
parameter_count = 0
byref_parameter_count = 0
local_variable_count = 0
call_fan_out = 0
```

An omitted or zero threshold is disabled. A positive threshold reports only
when the metric is greater than the threshold; negative values and invalid
types are configuration errors. Threshold evaluation emits metrics-specific
`MX001` warning entries, one entry for each exceeded procedure/metric pair.
`MX001` is not registered as a lint/analyze rule, cannot be inline-suppressed,
and is not affected by analyzer rule configuration. Without an enabled
threshold the command exits `0`. With one or more threshold findings it keeps
the complete metrics result, sets the failure error code to
`metrics_threshold_exceeded`, and exits `1`.

The v1 JSON contract is additive: consumers must ignore unknown fields and
additional metrics, while removal or a meaning change requires a new schema
version. Procedure arrays are sorted by normalized file, declaration start
position, module, name, and kind. Threshold diagnostics follow that procedure
order and the fixed metric order above. Paths, exclusion patterns, warnings, and JSON
serialization use normalized separators and stable ordering, so repeated runs
over the same source produce the same metrics and JSON bytes.

Parse recovery does not yield guessed or partial measurements. An unreadable
source, malformed exclusion pattern, or unrecoverable parse returns a metrics
operation error; an unmatched valid exclusion pattern is retained as a stable
warning. No generated file is silently counted twice.

## Consequences

- Metrics can be consumed by reports, baselines, and editor integrations
  without depending on diagnostic enablement or message wording.
- Threshold policy is explicit and can fail CI only when a project opts in;
  raw measurement collection remains useful with all thresholds disabled.
- The schema and ordering are stable across operating systems and source
  enumeration order, but consumers must honor `schema_version` before relying
  on field meaning.
- IR/CFG and call-resolution limitations remain visible: unresolved or
  ambiguous calls do not inflate fan-out, and unrecoverable syntax fails closed
  rather than producing misleading values.
- The procedure-complexity schema remains procedure-focused; architectural
  procedure/module hotspot ranking is an additive projection governed by
  ADR-0039. Historical trends, recommended thresholds, and LSP projections
  remain future work.

## Alternatives Considered

1. **Add complexity findings to `analyze`** - Rejected because diagnostics and
   measurements have different ownership, enablement, and compatibility needs;
   subjective thresholds would also affect existing `analyze` exit behavior.
2. **Reuse `analysis_metrics.module_state`** - Rejected because it is an
   analyzer-specific projection and does not contain the procedure structure,
   source-line semantics, or call fan-out required here.
3. **Compute metrics from diagnostic counts or formatted text** - Rejected
   because diagnostics are configurable and prose is not a deterministic
   source contract.
4. **Count unresolved and external calls as fan-out** - Rejected because an
   uncertainty is not evidence of a project-local dependency.
5. **Apply `[build].exclude` and analyzer rule settings automatically** -
   Rejected because release artifact selection and finding suppression are
   separate policies from maintainability measurement.

## Evidence

- Issue #460 acceptance criteria for deterministic metrics, configurable
  thresholds, JSON, exclusions, and syntax coverage.
- Procedure IR and source-backed declaration facts:
  `docs/specs/vba-analysis-ir.md` and ADR-0021.
- Conservative control-flow semantics:
  `docs/specs/vba-control-flow-graph.md` and ADR-0022.
- Existing project-local call resolution and deterministic ordering:
  `docs/specs/vba-call-graph-reachability.md`,
  `docs/specs/vba-procedure-call-cycles.md`, and ADR-0035.
- Public command and JSON envelope contract:
  `docs/specs/cli-contract.md` and `vitepress/reference/json-output.md`.
- Procedure metric definitions and threshold behavior:
  `docs/specs/vba-procedure-complexity-metrics.md`.

## Supersedes

None.

## Superseded by

None.

## Related

- Issue #460
- ADR-0021, ADR-0022, ADR-0035
- ADR-0039
- `docs/specs/vba-procedure-complexity-metrics.md`
