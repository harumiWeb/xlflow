# ADR-0034: SQL Execution Safety Ownership and Risk Taxonomy

## Status

Accepted

## Context

ADR-0025 established a conservative procedure-local source-to-sink layer and
`VBA224` currently exposes generic SQL execution flows. Issue #467 requires
SQL-specific handling for `CommandText`, `Recordset.Open`, parameterized ADO
commands, SQL roles, quoting, wildcard input, and locale-sensitive values.
Keeping these semantics entirely in `VBA224` would either lose the SQL context
or make the generic rule provider-specific. `VBA238` is reserved by parallel
development, so the next available public rule ID is `VBA239`.

## Decision

Add default-enabled, warning-level, procedure-local `VBA239`
(`detect_unsafe_sql_construction`) for recognized SQL execution boundaries.
Reuse ADR-0025's tainted/unknown provenance and CFG fixed point, while adding
small procedure-local database-object and `CommandText` states.

`VBA239` owns SQL findings when enabled. `VBA224` remains the compatibility
fallback when `VBA239` is disabled, and the two rules never emit duplicate SQL
findings for the same flow. `CommandText` is reported only when an execution on
the same receiver is reached in the same procedure. Parameter values bound
through recognized ADO parameter APIs are not treated as SQL text; dynamic
identifiers require a fixed allowlist because they cannot generally be bound.

The public context is additive: `sql_execution` carries the API, input role,
risk kind, origin state, and parameterized marker while retaining generic
`data_flow` provenance. Risk wording remains potential/conservative and does
not claim exploitability or complete injection proof.

## Consequences

- SQL users receive actionable value/identifier/LIKE/locale guidance without
  coupling all source-to-sink consumers to SQL semantics.
- Existing `VBA224` configurations continue to provide a fallback, but users
  must migrate SQL-specific suppressions to `VBA239` when both rules are enabled.
- Type/receiver recognition is intentionally explicit; unresolved or
  interprocedural construction remains outside v1 and may be reported as
  unknown only when the SQL boundary itself is established.
- New SQL observations can change corpus snapshots and require semantic review;
  snapshots remain observations rather than correctness approvals.

## Alternatives Considered

1. **Keep SQL under `VBA224` only.** Rejected because SQL API roles,
   parameterized commands, and identifier/value distinctions would make the
   generic contract provider-specific and difficult to consume.
2. **Report every dynamic SQL string as confirmed injection.** Rejected because
   unknown provenance and SQL correctness risks are not proof of exploitability.
3. **Require whole-project propagation in v1.** Rejected because ADR-0025's
   procedure-local boundary keeps batch and LSP behavior deterministic while
   summaries for ByRef and helper APIs remain unspecified.

## Evidence

- Issue #467 requirements for SQL sinks, parameterization, role distinctions,
  locale handling, and conservative wording.
- `docs/specs/vba-source-sink-dataflow.md` and
  `docs/adr/ADR-0025-conservative-vba-source-sink-dataflow.md`.
- `internal/vba/procedureir`, `internal/vba/dataflow`, and
  `internal/analyze/dataflow.go`.
- Existing corpus examples in `testdata/static-analysis-corpus/projects`,
  including ADO command and DAO dynamic SQL flows.

## Related

- Issue #467
- `docs/specs/vba-sql-execution-safety.md`
- `docs/adr/ADR-0024-shared-static-analysis-rule-registry.md`
- `docs/adr/ADR-0033-command-execution-safety-ownership.md`
