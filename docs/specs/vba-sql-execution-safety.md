# VBA SQL Execution Safety

<!-- xlflow-rule-contract: {"id":"VBA239","family":"analyze","category":"security","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_unsafe_sql_construction","inline_suppressible":true,"preflight_blocking":false} -->

`VBA239` is a default-enabled, warning-level, non-blocking diagnostic for
procedure-local SQL text assembled from external or unresolved input. It is
available in batch analysis and real-time diagnostics. The rule reports a
potential construction or injection risk; it never proves that a query is
exploitable or that an application is SQL-injection-free.

## Recognized execution boundaries

The first release recognizes SQL text passed to `ADODB.Connection.Execute`,
`ADODB.Recordset.Open`, DAO/Access `Database.Execute`, `CurrentDb.Execute`,
`OpenRecordset`, and `DoCmd.RunSQL`. For `ADODB.Command`, a `CommandText`
assignment is tracked only until an `Execute` on the same receiver in the same
procedure. DAO and ADO calls are recognized from explicit type evidence,
`New`/`CreateObject` origins, aliases, and known host receivers; arbitrary
members whose names merely resemble database APIs are not sinks.

`ADODB.Command.Execute` output arguments such as `recordsAffected` are not SQL
text. Constant-member `CallByName` forms are accepted only when the member and
invocation kind are statically known. Interprocedural construction and
execution are outside this contract.

## Safe and unsafe construction

Literal and `Const`-only SQL composition is clean. Values supplied only through
`CreateParameter`, `Parameters.Append`, or parameter `.Value` are accepted when
the executed `CommandText` remains constant. A dynamic `CommandText` remains a
finding even if other parameters are present.

Where the surrounding constant SQL makes the role clear, the analyzer labels
the input as a value, identifier, or `LIKE` pattern. It reports one most-specific
`risk_kind` per source/sink pair: `dynamic_identifier`,
`wildcard_like_input`, `locale_sensitive_value`, `manual_quoting`,
`external_value_concatenation`, or `unknown_origin`. Date and locale-sensitive
numeric conversions, manually quoted values, and wildcard-bearing `LIKE`
expressions receive targeted guidance. Identifiers should use a fixed allowlist;
values should use parameters and provider-appropriate wildcard escaping.

## Ownership and compatibility

When `VBA239` is enabled, SQL findings are projected only as `VBA239`; the
generic `VBA224` SQL flow is suppressed to avoid duplicates. Disabling `VBA239`
restores the `VBA224` fallback for compatibility. Existing `VBA224`
inline suppressions do not automatically suppress `VBA239` findings.

The additive `sql_execution` JSON context contains `risk_kind`, `api`,
`input_role`, `origin_state`, and `parameterized`, alongside the existing
`data_flow` source/sink/path context.

## Non-goals

- whole-project or interprocedural taint propagation;
- complete SQL parsing, provider-specific marker validation, or sanitizer proofs;
- parameterizing dynamic table/column identifiers;
- preflight blocking or Excel/VBE execution.
