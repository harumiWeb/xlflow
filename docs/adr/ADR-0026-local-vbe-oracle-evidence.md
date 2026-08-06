# ADR-0026: Local Excel/VBE Oracle Evidence Harness

## Status

Accepted

## Context

VBA has host-specific compile and call semantics that are not safely inferred
from general language rules. xlflow's lint, analyze, type-analysis, and LSP
diagnostics therefore need a way to check focused questions against a real VBE
without making Excel a dependency of normal analysis. A VBE answer is useful
evidence for compile-equivalent claims, but VBE acceptance does not invalidate
policy, safety, portability, or maintainability warnings.

There is no Excel-installed self-hosted GitHub Actions runner. Excel/VBE UI and
COM automation is also slow, stateful, and machine-dependent, so the check must
remain a developer-operated local validation tool.

## Decision

Add a Windows-only, developer-only oracle command backed by the existing .NET
Excel bridge and dialog infrastructure. The first schema supports only
standard-module compile probes. Go owns fixture loading and validation,
deterministic sequential selection, strict expectation checking, analyzer
contracts, and explicit promotion. The bridge owns one disposable workbook and
one real VBE Compile observation at a time.

Fixtures store source in ordinary `.bas` files and begin as `observe` with
pending provenance. An explicit promotion command may assert only confirmed
`accepted` or `rejected` outcomes for explicitly selected cases. Rejected
compile evidence is `compile-error`; accepted compile evidence must carry an
explicit non-runtime diagnostic meaning. Excel version/build, bitness, locale,
phase, and verification time are retained as provenance.

Only `accepted` and `rejected` are behavioral evidence. Timeouts, startup/VBOM
or import failures, missing Compile commands, unknown modals, worker/COM
failures, malformed output, and unconfirmed cleanup are infrastructure
failures. They stop a batch and cannot be promoted. Known accept/reject control
fixtures run first to detect a broken watcher or command invocation.

The oracle is never called by production xlflow commands, deterministic tests,
or GitHub Actions. Contributors changing static-analysis semantics must run
the relevant focused cases locally and record their IDs in the PR.

Binding metadata remains fixture-local. A rejected fixture may list accepted
`negative_controls`; Go validates the complete corpus and requires every bound
rule to have rejected expected evidence plus accepted forbidden evidence. The
referenced controls must cover all declared diagnostic surfaces, while
historical unbound fixtures are reported without becoming an initial CI gate.
This keeps the relationship reviewable beside the source that needs the
false-positive protection and makes rule-to-control coverage queryable without
launching Excel.

## Consequences

- Analyzer changes can cite empirical VBE compile evidence while normal tests
  remain deterministic and Excel-free.
- The observation/promotion boundary prevents guessed VBA behavior from being
  silently committed as authority.
- Local-only execution leaves a manual Windows prerequisite and cannot provide
  CI coverage until an Excel-capable self-hosted runner is available.
- Sequential disposable workbooks and strict cleanup handling reduce cross-case
  contamination at the cost of slower fixture runs.
- VBE evidence must remain separate from policy and maintainability rule
  ownership; accepted VBA may still produce xlflow warnings.
- Positive/negative coverage catches analyzer drift across future bindings while
  preserving the Excel-free CI boundary; the trade-off is that a fixture must
  remain `partially-bound` until every declared surface has an accepted control.

## Alternatives considered

1. **Invoke Excel from production lint/analyze/LSP** - Rejected because it
   would make diagnostics nondeterministic and require Office/COM everywhere.
2. **Use a PowerShell-only oracle implementation** - Rejected because the
   existing .NET bridge already owns Excel lifecycle, VBE Compile invocation,
   dialog correlation, timeout handling, and cleanup.
3. **Treat timeouts or missed dialogs as VBA rejection** - Rejected because
   infrastructure failure is not language evidence and would create false
   compile-error claims.
4. **Run fixtures in parallel** - Rejected because Excel/VBE command bars,
   modal dialogs, and process cleanup are single-user resources.
5. **Automatically promote observations** - Rejected because promotion is a
   reviewable evidence decision and must never overwrite an asserted result.

## Evidence

- `docs/specs/vbe-oracle.md`
- `docs/adr/ADR-0008-dotnet-excel-bridge.md`
- `docs/adr/ADR-0010-hybrid-excel-dialog-watcher.md`
- `docs/adr/ADR-0013-analyze-runtime-risk-ownership.md`
- `docs/adr/ADR-0024-shared-static-analysis-rule-registry.md`
- `bridge/dotnet/src/Xlflow.ExcelBridge/Services`
- `bridge/dotnet/src/Xlflow.ExcelBridge/Workers`
- `testdata/vbe-oracle/manifest.json`

## Related

- Issue #502
- Issue #514
- `docs/adr/ADR-0008-dotnet-excel-bridge.md`
- `docs/adr/ADR-0010-hybrid-excel-dialog-watcher.md`
- `docs/adr/ADR-0013-analyze-runtime-risk-ownership.md`
- `docs/adr/ADR-0024-shared-static-analysis-rule-registry.md`
