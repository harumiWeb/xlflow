# ADR-0033: VBA Error Outcomes and LSP Project Summary Cache

## Status

Accepted

## Context

ADR-0023 propagates possible procedure effects through uniquely resolved,
reachable project-local calls. Its `suppresses_errors` and `raises_error`
effects identify statements, but they do not describe whether an exceptional
path returns normally, rethrows, converts failure into a Boolean result, or
loses the original error in cleanup. Consequently, a caller cannot distinguish
an intentional `Try...`-style helper from a helper that silently hides failure.

Issue #446 requires `VBA237` to identify the procedure or call boundary where
failure information is lost, including multi-level call chains. The analysis
must remain high precision: ambiguous, unresolved, external, and dynamic calls
are uncertainty rather than proof, and an error handler is not unsafe merely
because it exists.

The LSP adds a second architectural concern. Interprocedural results require a
coherent project summary, while ADR-0014 permits open-document overlays to be
temporarily pending and publishes procedure-local Fast diagnostics before the
Full pass. Recomputing and republishing every open document after every edit
would be correct but unnecessarily expensive; reusing a summary across an
affected call edge could publish stale failure chains.

## Decision

Extend `internal/vba/effects.ProcedureSummary` with a provenance-bearing error
outcome summary derived from procedure IR, CFG paths, and normalized error
statements. The summary distinguishes handler presence, resume-next use,
suppression, explicit rethrow, Boolean success return, possible raise, and
logging followed by normal continuation. These are finite facts with source
evidence, not independent syntax flags: conclusions about suppression and
rethrow are based on reachable error-to-exit outcomes.

Propagate error outcomes and their provenance by deterministic set union over
the same uniquely resolved, reachable project-local call edges established by
ADR-0023. Ambiguous, unresolved, external, and dynamic calls retain explicit
uncertainty and cannot independently produce `VBA237`. Fixed-point propagation
must converge for recursion, mutual recursion, and diamond-shaped call graphs,
and representative chains must remain deterministic.

Make the failure-loss boundary the diagnostic owner. A handler or cleanup path
that returns normally without a rethrow, explicit fallback, or failure result
is reported at that handler or cleanup boundary. A caller that discards a
known Boolean success result is reported at the call site. Public entry points
and host events may be named as representative context, but do not receive
duplicate ancestor diagnostics. Intentional checked probes, explicit
`Err.Raise` or `Error` rethrows, explicit recovery, fallback-return functions,
and checked `Try...`-style Boolean helpers remain accepted.

Cache an immutable project error summary for each complete LSP workspace
revision. Workspace entries retain only Go-owned procedure IR and CFG values;
the cache does not retain tree-sitter nodes or LSP protocol values. Open buffers
remain authoritative. While their newest overlay is pending, do not reuse the
saved or previous overlay summary and do not publish `VBA237`; rebuild after
the matching overlay is complete.

`VBA237` is a Full-only LSP diagnostic. Fast analysis neither computes,
rebases, nor reuses it. Project-summary updates compare direct outcome
fingerprints, call-resolution status, uncertainty, and confirmed edges. Find
affected procedures by traversing the union of the old and new confirmed call
graphs in reverse, so both added and removed dependencies invalidate their
transitive callers. Schedule Full diagnostics only for affected open VBA
documents.

Track a dependency generation for each affected document. A Full result may
publish only if its document lifecycle, snapshot generation, project-summary
revision, and dependency generation still match. Overlay publication,
`didClose` restoration, watched-file changes, and file deletion use the same
reverse-dependency invalidation. Canceled, incomplete, or obsolete summaries
are not cached as completed values.

Expose no new summary fields through CLI JSON or LSP. `VBA237` is a
default-enabled, high-precision, interprocedural warning on the `analyze` and
LSP surfaces; it is inline-suppressible, configurable with
`detect_error_suppression_propagation`, and non-blocking for preflight.

## Consequences

- Positive: direct and multi-level failure suppression can be explained with
  stable provenance while intentional handlers and checked helpers remain
  valid.
- Positive: diagnostics point to the actionable boundary where failure
  information disappears instead of repeating one root cause at every caller.
- Positive: unresolved dispatch remains visible as uncertainty and cannot be
  converted into a high-confidence failure claim.
- Positive: Full LSP diagnostics remain coherent with unsaved overlays, and
  changes refresh only open documents that can depend on the changed summary.
- Negative: outcome summaries require path-sensitive CFG queries in addition
  to the existing possible-effect vocabulary.
- Negative: retaining old and new reverse edges during replacement and
  tracking per-document dependency generations add LSP cache complexity.
- Negative: Full-only publication delays `VBA237` after an edit until the
  workspace overlay and project summary are complete.
- Limitation: passing a success flag as an argument or calling external or
  dynamically dispatched code remains uncertain unless a future contract can
  prove how that value or failure is observed.

## Alternatives Considered

1. **Treat every error handler or `Resume Next` as suppression** - Rejected
   because explicit rethrows, recovery, narrow checked probes, and intentional
   fallback helpers are valid VBA error-handling patterns.
2. **Propagate one Boolean `may_suppress` flag** - Rejected because callers
   could not explain where information was lost, distinguish outcomes, or
   deduplicate recursion and diamond paths reliably.
3. **Report public entry points that cannot observe a failure** - Rejected as
   the primary location because every ancestor would repeat the same loss and
   obscure the handler or ignored result that should be fixed.
4. **Run the rule during Fast diagnostics** - Rejected because Fast analysis
   deliberately lacks a fully revalidated project call graph and cannot safely
   rebase interprocedural evidence.
5. **Invalidate every open document after any project edit** - Rejected because
   reverse-call dependencies provide the same safety with substantially less
   repeated Full analysis.
6. **Traverse only the new reverse call graph** - Rejected because removing or
   making an old edge ambiguous must also invalidate callers that depended on
   the previous summary.

## Evidence

- Requirements: xlflow issue #446.
- Existing procedure summaries and propagation contract:
  `docs/adr/ADR-0023-procedure-effect-summaries.md` and
  `docs/specs/vba-effect-analysis.md`.
- Conservative exceptional flow and non-returning raise semantics:
  `docs/adr/ADR-0022-conservative-vba-control-flow-graph.md` and
  `docs/specs/vba-control-flow-graph.md`.
- LSP snapshot, overlay, Fast/Full, and stale-publication contract:
  `docs/adr/ADR-0014-reusable-vba-lsp-server.md` and
  `docs/specs/cli-contract.md`.
- Detailed outcome, diagnostic, and invalidation contract:
  `docs/specs/vba-error-propagation-analysis.md`.

## Supersedes

None.

## Superseded by

None.

## Amended by

- `docs/adr/ADR-0044-bounded-procedure-effect-summaries.md`

## Related

- Issue #446
- `docs/adr/ADR-0014-reusable-vba-lsp-server.md`
- `docs/adr/ADR-0021-procedure-analysis-ir.md`
- `docs/adr/ADR-0022-conservative-vba-control-flow-graph.md`
- `docs/adr/ADR-0023-procedure-effect-summaries.md`
- `docs/specs/vba-error-propagation-analysis.md`
