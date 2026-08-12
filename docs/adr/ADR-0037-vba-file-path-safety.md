# ADR-0037: VBA File and Path Safety Ownership

## Status

Accepted

## Context

ADR-0025 introduced generic procedure-local source-to-sink dataflow and
`VBA224`, which can identify some tainted values reaching file and workbook
sinks. Issue #468 also requires findings for dangerous constants and path
semantics such as relative paths, wildcards, traversal, overwrite behavior,
same-source/destination operations, and temporary-file cleanup. Those facts
cannot be represented by taint alone.

## Decision

Add default-enabled, warning-level, procedure-local `VBA245`
(`detect_unsafe_file_path`) for destructive and state-dependent file/path
operations. It is available in batch and realtime/LSP analysis, is inline
suppressible, and does not block preflight. The rule owns recognized VBA file
statements, FileSystemObject write/delete/copy/move methods, and workbook
`SaveAs`/`SaveCopyAs` path findings.

`VBA245` performs conservative local path construction tracking for literals,
aliases, concatenation, `BuildPath`, trusted symbolic anchors such as
`ThisWorkbook.Path` and configured project paths, and the existing external
input catalog. It classifies definite path hazards separately from
input-dependent and temporary-lifecycle risks in an additive
`file_operation` context. Existence checks are not sinks. Runtime filesystem
identity, symlinks, junctions, and interprocedural cleanup summaries remain out
of scope.

When enabled, `VBA245` owns the overlapping file/workbook sinks and suppresses
the corresponding generic `VBA224` projection. When disabled, `VBA224` remains
the compatibility fallback for its existing destructive-file and SaveAs
source-to-sink coverage; non-destructive workbook-open flows remain under
`VBA224` regardless of this setting.

## Consequences

- Constant path hazards are visible even without a tainted source.
- Clients receive stable operation, path-role, risk, origin, anchor, and
  overwrite metadata rather than parsing prose.
- Existing `VBA224` suppressions for file sinks may need migration when the
  specialized rule is enabled.
- The first lifecycle analysis is procedure-local and may conservatively warn
  when cleanup cannot be proven in the same procedure.

## Alternatives considered

1. Keep path policy in `VBA224`. Rejected because clean dangerous constants and
   operation-specific overwrite/path roles would make the generic contract
   provider-specific.
2. Create separate IDs for every file operation. Rejected because it would
   fragment suppression and duplicate the shared path analysis.
3. Treat every dynamic path as a confirmed vulnerability. Rejected because
   unknown origin and runtime policy are not exploitability proof.

## Evidence

- Issue #468 requirements.
- `docs/specs/vba-source-sink-dataflow.md` and ADR-0025.
- `internal/vba/dataflow` and `internal/analyze/dataflow.go`.
- `docs/specs/vba-file-path-safety.md` and focused `VBA245` tests.

## Related

- Issue #468
- ADR-0025
- ADR-0033
- ADR-0034
