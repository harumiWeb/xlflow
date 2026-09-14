# ADR-0052: Procedure-Local Active-State Safety for Excel Select Operations

## Status

Accepted

## Context

`VB003` intentionally discourages `Select` and `Activate`, but it cannot tell
whether an intentional UI operation has the Excel active-state preconditions
required by the object model. `Worksheet.Select` requires the parent workbook
to be active, and `Range.Select` requires the owning worksheet to be active.
Refactoring code to remove intermediate activation calls can therefore leave a
valid VBA procedure that fails only when it runs against a different active
workbook or worksheet.

The problem is path-sensitive and must remain separate from the style policy.
The analyzer already has procedure-local CFG and semantic facts, but it must
not infer UI state from a helper procedure or from an uncertain external call.

## Decision

Add `VBA250` as a default-enabled, warning-level, non-blocking, inline-
suppressible `runtime-safety` analyzer rule. It is procedure-local and is
projected by batch `analyze` and realtime/LSP diagnostics. Its canonical
configuration key is `detect_unsafe_select_operations`; the shared
`[analyze].disabled_rules` policy remains authoritative when both forms are
present.

The rule uses a small abstract state with independent workbook and worksheet
identities. Both facts are unknown at procedure entry. A join keeps only an
identity proven on every reachable incoming path. `Workbook.Activate` sets the
workbook fact and clears the worksheet fact; `Worksheet.Activate` sets both;
`Worksheet.Select` requires the parent workbook fact; and `Range.Select`
requires the owning worksheet fact. Uncertain edges, `On Error Resume Next`,
unknown calls, and unsupported state-changing APIs invalidate the relevant
facts. A helper's activation is never imported into its caller in this initial
scope.

Only resolved Excel object-model references participate. Common local aliases,
`ThisWorkbook`, worksheet collections, range roots, and `With` blocks may be
normalized when identity is available. Non-Excel same-named members,
late-bound receivers, and unresolved aliases remain unknown and do not create a
false proof of safety.

`VBA250` owns runtime precondition findings only. `VB003` continues to report
the independent maintainability concern that a selection or activation may be
avoidable. Neither rule changes the other's enablement or suppression policy.

## Consequences

Positive consequences:

- Intentional result-sheet navigation remains valid while missing activation
  preconditions become visible before runtime.
- Conditional activation and branch joins are handled conservatively without
  claiming that valid VBA is rejected by VBE.
- Batch and realtime projections share one registry contract and one
  procedure-local state model.

Negative consequences:

- Calls whose state effects are not modeled remain unknown and may produce no
  finding even when Excel would fail.
- The first version does not propagate active state across procedure calls or
  model every Excel navigation API.
- A new rule ID and compatibility key must be carried through registry,
  configuration, generated references, and public CLI documentation.

## Alternatives considered

1. **Extend `VB003` with active-state reasoning.** Rejected because style
   guidance and runtime preconditions have different evidence, configuration,
   and suppression semantics.
2. **Assume every `Activate` or helper call establishes state globally.**
   Rejected because conditional branches, unknown calls, and procedure-local
   scope would create false negatives.
3. **Use Excel/VBE during normal analysis.** Rejected because production
   diagnostics remain source-only, deterministic, and parallel-safe.
4. **Report every `Select` without proving failure.** Rejected because it
   would duplicate `VB003` and turn intentional UI navigation into noise.

## Related

- Issue #792
- `docs/adr/ADR-0024-shared-static-analysis-rule-registry.md`
- `docs/adr/ADR-0046-procedure-applicability-planning.md`
- `docs/specs/vba-unsafe-select-diagnostics.md`
