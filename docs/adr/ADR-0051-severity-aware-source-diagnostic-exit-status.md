# ADR-0051: Severity-Aware Source Diagnostic Exit Status

## Status

`accepted`

## Background

`lint` and `analyze` already expose a severity for every source diagnostic,
but their CLI handlers previously returned validation exit code `1` whenever
the result contained any issue or finding. That made advisory warnings look
like command failures to AI agents and automation even though the diagnostics
remain useful output and do not indicate a compile-equivalent error.

The `check` command combines the same source diagnostics with the independent
Excel environment check, so it must apply the same source-diagnostic policy to
avoid presenting a different result for the same project.

## Decision

Treat `error`-severity source diagnostics, and diagnostics with unknown or empty
severity, as validation failures for the `lint`, `analyze`, and
source-diagnostic portion of `check` commands. Warning and information
diagnostics remain in the JSON and human-readable output, but the command
reports status `ok` and returns exit code `0` when no blocking diagnostic is
present.

The fail-closed treatment of unknown or empty severity prevents a malformed
diagnostic from silently becoming a successful result. Configuration,
analysis/lint execution, and Excel environment failures retain their existing
exit-code classes. The diagnostic payload and ordering are unchanged.

## Consequences

- AI agents and CI can distinguish advisory source feedback from a failed
  validation without losing warning details.
- Error-level diagnostics continue to fail the command and remain visible.
- Consumers that intentionally treated every diagnostic as a failure must
  inspect severity or the returned status and update their policy.
- `check` remains able to return exit code `3` for an Excel/operational
  failure, even when source diagnostics are warning-only.

## Alternatives Considered

1. **Keep counting every diagnostic as a failure.** Rejected because warning-
   only results are advisory and cause unnecessary agent retries or failure
   branches.
2. **Use registry `preflight_blocking` metadata as the CLI exit policy.**
   Rejected because it describes whether `push`/`run` may proceed, not whether
   a batch diagnostic is an error-level result for the caller.
3. **Hide or suppress warning diagnostics.** Rejected because agents and users
   still need the source locations and remediation guidance.

## Evidence

- CLI handlers: `internal/cli/root.go` (`lintCommand`, `analyzeCommand`, and
  `checkCommand`).
- Severity models: `internal/lint/linter.go` and
  `internal/analyze/analyzer.go`.
- Regression coverage: `internal/cli/root_test.go` warning-only and
  error-level command tests.
- Public contract: `docs/specs/cli-contract.md`, the README files, and the
  VitePress `lint` and `analyze` command pages.

## Supersedes

- None.

## Superseded by

- None.

## Related

- `docs/adr/ADR-0013-analyze-runtime-risk-ownership.md`
- `docs/adr/ADR-0049-rooted-headless-gui-preflight.md`
