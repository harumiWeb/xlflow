# ADR-0025: Conservative VBA Source-to-Sink Dataflow

## Status

Accepted

## Context

VBA source can receive values from workbook cells, files, environment state,
HTTP, databases, and user input before passing them to process, file, workbook,
SQL, or HTTP APIs. Existing analyzer rules identify individual runtime-risk
patterns but do not preserve provenance from an untrusted source to a sensitive
sink.

The first implementation must remain useful in batch analysis and the LSP
without making unsupported VBA semantics look safe. ProcedureIR and the
conservative procedure CFG already provide the normalized expressions,
statement ranges, reachability, and uncertain control-flow edges needed for a
procedure-local analysis.

## Decision

Add a protocol-neutral `internal/vba/dataflow` layer above `procedureir` and
`cfg`. It performs a forward, procedure-local may-analysis with `clean`,
`tainted`, and `unknown` states. Parameters and the initial source catalog are
tainted at their producing expressions. Direct assignment, local aliases, and
string concatenation preserve provenance. Unsupported or recovered
transformations remain unknown and are reported conservatively at a sink.

The initial sink catalog is explicit rather than matching arbitrary member
names. It covers process launch, SQL execution, destructive file operations,
workbook open/save, and HTTP URLs/headers. The source catalog covers parameters,
worksheet cell values, `InputBox`, text/CSV input, environment variables,
command-line arguments, HTTP responses, and database results.

V1 does not propagate taint across procedures. A resolved call without an
available taint summary is an unknown transformation. Future summaries may be
added without changing the procedure-local transfer contract.

Only narrow safety contracts clear taint: literal/constant-only values,
HTTP-specific `EncodeURL`, and a proven constant allowlist branch. General
`Trim`, `CStr`, `Replace`, `IsNumeric`, and `Len` calls are not considered
sanitizers by themselves.

Expose the result as default-enabled, non-blocking warning `VBA224`. Findings
retain source, sink, and a deterministic propagation path in the additive
`data_flow` JSON context, and include the conservative-analysis label in human
and LSP messages. Existing suppression and rule-registry policies remain the
source of truth for configuration and editor behavior.

## Consequences

- Batch and LSP analysis share one provenance implementation.
- Unknown syntax and unsupported transformations cannot silently become safe.
- The explicit catalogs limit false positives from unrelated `.Run`, `.Execute`,
  or `.Value` members.
- Initial coverage intentionally misses interprocedural flows and user-defined
  sanitizer summaries until a separate summary contract is designed.
- `VBA224` adds an additive analyzer JSON field and a new registry entry that
  documentation and editor consumers must project consistently.

## Alternatives Considered

1. **Regex-only sink checks** - Rejected because aliases, concatenation,
   branch guards, and provenance paths would drift between batch and LSP.
2. **Treat every unknown call as clean** - Rejected because it would create
   unsound false negatives at security-sensitive sinks.
3. **Propagate through every project call in V1** - Rejected because unresolved,
   ambiguous, ByRef, and return-value semantics need summaries before they can
   be trusted.
4. **Treat generic string helpers as sanitizers** - Rejected because trimming,
   quoting, or replacement is not a sink-independent security proof.

## Evidence

- `internal/vba/procedureir` and `docs/specs/vba-analysis-ir.md`
- `internal/vba/cfg` and `docs/specs/vba-control-flow-graph.md`
- `internal/analyze/analyzer.go`
- `internal/staticanalysis/rules` and `docs/adr/ADR-0024-shared-static-analysis-rule-registry.md`

## Related

- Issue #465
- `docs/specs/vba-source-sink-dataflow.md`
- `docs/adr/ADR-0021-procedure-analysis-ir.md`
- `docs/adr/ADR-0022-conservative-vba-control-flow-graph.md`
