# ADR-0032: Reviewed Static-Analysis Corpus Evidence

## Status

Accepted

## Context

ADR-0031 makes real-world `lint` and `analyze` output deterministic and
reviewable as golden snapshots. Those snapshots answer whether analyzer output
changed, but a snapshot row remains an observation: it does not say whether a
diagnostic is correct. Regenerating a snapshot can therefore preserve a false
positive indefinitely, while treating every snapshot row as approved would
conflate regression stability with semantic quality.

Issue #537 requires a second evidence layer that records reviewed true and
false positives, keeps unreviewed findings visible, and produces reproducible
quality metrics. The layer must retain known false-positive boundaries after
the analyzer is fixed so the diagnostic cannot silently reappear. It must also
reuse the expected/forbidden diagnostic contract concepts established for the
local VBE oracle without making corpus verification depend on Excel or on the
oracle runner.

## Decision

Store reviewed evidence in one UTF-8 JSONL ledger at
`testdata/static-analysis-corpus/reviews/diagnostics.jsonl`. Each row has
`schema_version`, stable `project` and project-relative `file` identities, a
`classification` of `true-positive` or `false-positive`, a diagnostic contract
containing canonical `code`, `severity`, `surface`, and the complete normalized
`range`, a positive `count`, and a non-empty `rationale`. False-positive rows
also contain exactly one of `regression_test` or `regression_exception`. When
other reviewed findings share a false-positive row's complete normalized
identity, `allowed_occurrences` records their current ceiling without changing
the false-positive evidence count.

The complete four-coordinate range is part of identity. It uses the same
normalized source positions as the production corpus result, including the
deterministic fallback from a missing end position to the start position. A
review therefore applies only to the exact project, file, surface, rule,
severity, and range recorded in the ledger; it cannot accidentally approve or
forbid a different occurrence of the same rule elsewhere.

A true-positive row is an expected diagnostic contract: the current corpus
must contain the exact diagnostic at least the declared number of times.
Additional occurrences remain unreviewed. A false-positive row is a forbidden
diagnostic contract: after remediation the current corpus must not add an
occurrence of that identity. The ordinary case has a zero ceiling;
`allowed_occurrences` handles the narrower case where
multiple source flows collapse to the same normalized identity and fails on
any increase above the reviewed baseline. The false-positive row remains
committed after remediation, preserving both the regression boundary and the
reviewed evidence used by quality metrics.

Diagnostics in current snapshots that have no matching ledger row are
`unreviewed`. They remain permitted and visible and are not silently promoted
to true positives. Ledger rows are validated and sorted deterministically;
conflicting or ambiguous classifications for the same exact identity are
invalid.

Quality summaries are derived from the committed ledger and current corpus
observations. They report reviewed, unreviewed, true-positive, and
false-positive counts and per-rule precision, with optional project/profile
coverage. `count` contributes diagnostic multiplicity. Precision is
`TP / (TP + FP)` and excludes unreviewed observations. Retaining remediated
false-positive rows prevents a fix from erasing the historical evidence and
artificially improving the metric.

False-positive remediation starts from the real-world occurrence, reduces it
to a minimal focused fixture where practical, fixes the root cause, adds the
focused regression test, and then commits or retains the real-world forbidden
contract. `regression_test` identifies that focused protection. A
`regression_exception` is allowed only when a focused fixture is impractical
and must explain why; it is not a shortcut around the normal regression
requirement.

The expected/forbidden matcher and diagnostic contract model belong to shared
static-analysis infrastructure. The corpus and VBE oracle may consume that
generic layer, but neither owns the other's evidence. VBE compile acceptance
or rejection can support compile-equivalent rules only. Successful compilation
does not prove that maintainability, policy, runtime-risk, portability, or
Excel-specific warnings are false positives.

## Consequences

- Snapshot stability and reviewed correctness become separate, composable
  evidence layers.
- Known false positives become durable forbidden contracts instead of rows
  that disappear from both snapshots and institutional memory after a fix.
- Unreviewed findings can accumulate without blocking the corpus while their
  exclusion from precision prevents unsupported quality claims.
- Exact ranges and multiplicity make matching deterministic, but source or
  diagnostic-range movement requires explicit evidence review.
- `allowed_occurrences` detects an increase in a colliding normalized identity,
  but cannot distinguish one allowed occurrence disappearing while one false
  positive reappears. That distinction would require adding rule-specific
  source/path context to the shared corpus identity.
- Reviewers must maintain rationales and focused fixture references; justified
  exceptions remain visible rather than silently weakening regression
  coverage.
- Shared contract code reduces semantic drift between corpus and oracle
  checks, while the evidence sources and authority boundaries remain distinct.

## Alternatives Considered

1. **Treat every snapshot diagnostic as a true positive** - Rejected because a
   stable snapshot proves repeatability, not correctness.
2. **Delete false-positive evidence after the fix** - Rejected because the
   diagnostic could reappear and historical precision would change merely
   because remediation succeeded.
3. **Match only by rule code and start line** - Rejected because it can bind a
   review to another surface, severity, occurrence, or source span.
4. **Require every current diagnostic to be reviewed immediately** - Rejected
   because broad corpus coverage should remain useful while semantic review is
   incremental.
5. **Use VBE compilation as the universal false-positive oracle** - Rejected
   because many xlflow rules intentionally diagnose compiling code.
6. **Keep corpus contracts inside the VBE oracle package** - Rejected because
   deterministic corpus validation is Excel-free static-analysis
   infrastructure and must not depend conceptually on the local oracle runner.

## Evidence

- Issue #537 requirements for reviewed expected/forbidden contracts,
  false-positive remediation, and reviewed-only metrics.
- Corpus snapshots and operational workflow:
  `docs/specs/static-analysis-corpus.md`.
- Deterministic snapshot identity and update boundary:
  `docs/adr/ADR-0031-deterministic-static-analysis-snapshots.md`.
- VBE evidence authority and compile-equivalent limitation:
  `docs/adr/ADR-0026-local-vbe-oracle-evidence.md`.
- Shared matcher and corpus enforcement:
  `internal/staticanalysis/contract/contract.go` and
  `internal/staticanalysis/corpus/review.go`.
- Deterministic contract and metrics tests:
  `internal/staticanalysis/contract/contract_test.go` and
  `internal/staticanalysis/corpus/review_test.go`.
- Committed reviewed evidence:
  `testdata/static-analysis-corpus/reviews/diagnostics.jsonl`.

## Supersedes

None.

## Superseded by

None.

## Related

- Issue #537
- `docs/specs/static-analysis-corpus.md`
- `docs/adr/ADR-0026-local-vbe-oracle-evidence.md`
- `docs/adr/ADR-0031-deterministic-static-analysis-snapshots.md`
