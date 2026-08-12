# ADR-0038: Project-Level Preflight Diagnostic Waivers

## Status

Accepted

## Context

ADR-0024 established the canonical static-analysis registry and made
compile-equivalent rules source-preflight blockers. That default protects
automation from modal Excel/VBE compile failures, but a false positive can
prevent every workbook-writing command even when a project owner has reviewed
and accepted the risk.

Adding rule-specific eligibility flags or separate lint/analyze allowlists
would create another policy layer for users and maintainers. The registry
already identifies the complete set of diagnostics that participate in source
preflight.

## Decision

Add `[preflight].allowed_diagnostics` as a project-level list of diagnostic
codes. Its default is empty. Configuration loading trims and uppercases codes,
ignores empty entries, removes duplicates while preserving first occurrence,
and rejects unknown codes or registry rules whose `preflight_blocking` value is
false.

Every registry rule with `preflight_blocking = true` is uniformly eligible.
The registry value continues to describe the default policy exposed by
`xlflow rules`; no second per-rule waiver-eligibility field or family-specific
allowlist is introduced.

The shared source-preflight policy classifies a registry diagnostic as
non-blocking, blocking, or explicitly allowed. CLI source-preflight consumers
use that policy for lint-owned and analyzer-owned findings. Applied waivers are
deduplicated by source occurrence, aggregated by diagnostic code, and preserved
as structured command warnings even when another blocker or a later Excel/VBE
compile step fails.

A waiver changes only whether the shared source-preflight gate stops workbook
automation. It does not disable the diagnostic, alter its severity, change
`lint` or `analyze` exit behavior, affect LSP projection, or permit inline
suppression. Parser/analysis execution failures, unreadable source, duplicate
components, and UserForm `FRM...` / `UFY...` integrity failures have no registry
diagnostic ID and remain non-waivable.

## Consequences

- Existing projects retain the current fail-closed default.
- Project owners can accept a known diagnostic occurrence without hiding it
  from static-analysis or editor surfaces.
- A waived source-preflight finding can still fail VBE compilation, so every
  applied waiver produces an explicit warning.
- New preflight-blocking registry diagnostics automatically participate in the
  same project policy without a parallel eligibility update.

## Alternatives considered

1. Allow only selected compile-equivalent rules. Rejected because a second
   eligibility taxonomy increases cognitive cost and can drift from the
   canonical registry.
2. Add separate lint and analyzer allowlists. Rejected because source
   preflight is one cross-family policy and diagnostics already have unique
   canonical IDs.
3. Disable or lower the severity of waived diagnostics. Rejected because the
   waiver accepts automation risk; it does not invalidate the diagnostic.
4. Add a wildcard or CLI bypass flag. Rejected because waivers should be
   explicit, reviewable project policy and should not silently cover future
   blockers.

## Evidence

- `internal/staticanalysis/preflight` owns the protocol-neutral policy.
- Config tests validate normalization and every registry blocker.
- CLI tests cover lint/analyze blockers, deterministic warning aggregation,
  repeated preflight, mixed failures, and unchanged batch diagnostics.
- Generated configuration and diagnostic documentation describe the default
  registry policy and project waiver boundary.

## Amends

- ADR-0024

## Related

- Issue #616
- ADR-0036 (resolution diagnostics)
- `docs/specs/cli-contract.md`
- `vitepress/reference/config-file.md`
