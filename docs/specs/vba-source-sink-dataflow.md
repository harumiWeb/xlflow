# Conservative VBA Source-to-Sink Dataflow

This specification defines the generic procedure-local dataflow contract for
`VBA224`. Destructive/path-specific classification and diagnostics belong to
`VBA245`; process-launch-specific classification and diagnostics belong to
`VBA236`, and SQL-specific classification belongs to `VBA239`. `VBA224`,
`VBA236`, and `VBA239` reuse the generic `procedureir.ProcedureIR` expression
states and CFG paths through `dataFlowFindingsContext`; `VBA245` is invoked by
the same batch and realtime procedure loops but performs its own lexical,
procedure-local path pass over source lines and parsed declarations. It does
not project through the generic sink findings, which keeps its structured
file-operation metadata and ownership separate.
The implementation is protocol-neutral and consumes `procedureir.ProcedureIR`
and one `cfg.Graph`; it does not parse source or depend on CLI/LSP types.

## State and transfer

The analysis uses `clean`, `tainted`, and `unknown` facts. `tainted` and
`unknown` are unsafe at a sensitive sink. Facts are merged by may-analysis, so
one unsafe incoming path keeps the result unsafe after a CFG join. Exceptional,
termination, and unknown CFG edges remain conservative inputs.

The following operations preserve provenance:

- procedure parameters start tainted;
- assignment and simple variable aliases copy the source fact;
- string concatenation is unsafe when any operand is unsafe;
- unknown or recovered transformations retain an `unknown` step when they
  consume an unsafe value;
- procedure calls do not propagate taint across a procedure boundary in V1.

The analyzer uses expression IDs and statement links from ProcedureIR rather
than rescanning comments or string literals. Unsupported expression shapes are
unknown and never prove a value safe.

## Initial sources

The source catalog recognizes:

- procedure parameters;
- `Range`/`Cells` worksheet value reads;
- `InputBox` and `Application.InputBox` results;
- VBA text/CSV input (`Line Input`, `Input #`) and recognized text-file reads;
- `Environ`/`Command$` and recognized WScript argument/environment reads;
- recognized HTTP response members such as `ResponseText` and `ResponseBody`;
- recognized ADO recordset/field result reads.

Matching is case-insensitive and limited to the explicit catalog. Arbitrary
members with names such as `Value` or `ResponseText` are not automatically
sources without the required receiver shape or type evidence.

## Initial sinks

The shared data-flow catalog also carries process-launch and SQL entries for
the `VBA236`/`VBA239` projections; the generic `VBA224` finding set uses only
the entries not owned by an enabled specialized rule.

The sink catalog recognizes:

- recognized ADO SQL execution arguments;
- legacy destructive file and SaveAs paths as a compatibility fallback when
  `VBA245` is disabled (workbook-open remains a generic non-destructive sink);
- HTTP request URLs and `setRequestHeader` values for recognized HTTP clients.

Generic `.Run`, `.Execute`, or `.Open` members do not become sinks without an
explicit catalog match; `SaveAs` is the explicit workbook-save sink contract.
Process-launch sinks are intentionally excluded from `VBA224`; `VBA236` owns
VBA `Shell`, `WScript.Shell.Run`/`Exec`, `Shell.Application.ShellExecute`, and
Win32 `ShellExecute` variants so executable paths, interpreter command text,
arguments, URL/document targets, and observation state can be classified
without duplicate diagnostics.

## Safety contracts and guards

Literal and constant-only expressions are clean. Constant identifiers include
local and current-module `Const` declarations plus host constants resolved by
the shared VBA type database, such as `vbCrLf` and `vbNullString`. A local or
module variable with the same name takes precedence and remains subject to the
conservative variable-state rules. The initial sink-specific contracts
additionally recognize `EncodeURL` for HTTP URLs and a proven constant allowlist
expressed by exact equality or `Select Case`. A guard only applies on the CFG
branch where it is true; an unsafe alternative branch keeps the value unsafe
after the join.

`Trim`, `CStr`, generic `Replace`, `IsNumeric`, and `Len` do not clear taint.
User-defined sanitizer functions and interprocedural validation summaries are
out of scope for V1.

## Diagnostic contract

`VBA224` is a default-enabled, procedure-local, realtime warning in the
`security` category. It is inline-suppressible and does not block preflight.
Its optional `data_flow` JSON context contains one source, one sink, and a
deterministic representative path:

```json
{
  "source": { "kind": "inputbox", "label": "InputBox", "line": 3 },
  "sink": { "kind": "sql_execution", "label": "SQL execution", "line": 5 },
  "path": [
    { "kind": "assignment", "label": "raw = InputBox(...)", "line": 3 },
    { "kind": "concatenation", "label": "command = prefix & raw", "line": 4 }
  ]
}
```

The message and reason identify the source, sink, path, and that the result is
conservative. Findings are deduplicated per source/sink pair and sorted by the
existing analyzer ordering contract. Fixed-point propagation completes before
finding emission. Each reachable statement block is then evaluated once from
its converged block-entry state, so transient states and worklist priority
cannot add or remove diagnostics. Changing the fixed-point traversal order
must preserve the finding set, state, source/sink identity, and representative
path.

When the sink is a process launch, this generic context is projected by
`VBA236` into its additive `command_execution` object; `VBA224` does not emit a
second process-launch finding.

When `VBA245` is enabled, its file-operation projection owns the destructive
file and workbook-save sinks, including clean-but-dangerous constants; the
generic `VBA224` projection is suppressed for those sinks. Disabling `VBA245`
restores the legacy `VBA224` fallback.

## Non-goals

- interprocedural taint propagation or user-defined sanitizer summaries;
- whole-project aliasing or precise COM type inference;
- preflight blocking, automatic fixes, or Excel COM execution;
- treating every unknown call/member as a source or sink.
