# ADR-0041: Specific VBA Syntax Recovery Precedence

## Status

Accepted

## Context

The tree-sitter VBA grammar intentionally recovers malformed source. The lint
pipeline therefore needs `VB014` as a compatibility fallback, but several
recovery shapes have a high-confidence interpretation shared by lint and LSP.
The previous precedence rule suppressed every `VB014` in a file whenever an
unrelated syntax diagnostic existed, hiding actionable recovery elsewhere and
making colon-separated statements indistinguishable.

## Decision

Keep parser recovery as a first-class internal range and apply precedence per
recovery range:

1. A high-confidence syntax diagnostic owns only the overlapping `ERROR` or
   `MISSING` range.
2. The overlapping generic `VB014` is suppressed.
3. Any recovery range without a specific owner remains `VB014`.

Diagnostics that do not expose a source span retain a narrow same-line
fallback for compatibility. Structural recovery with no parser range is only
discarded when a specialized recovery rule owns the otherwise location-less
finding; unrelated parser recovery with an explicit range is retained.

Recovery interpretation is centralized in lint helpers for conditional branch
ordering (`VB062`), Select/Case ordering (`VB063`), Open mode recovery
(`VB064`), and the unambiguous trailing-token TypeOf shape (`VB065`).
ERROR-only or keyword-shifted forms that cannot be distinguished reliably stay
under `VB014`. Existing ownership remains unchanged for `VB009`, `VB010`-
`VB013`, `VB047`-`VB049`, `VB059`, and array diagnostics.

## Consequences

- A single malformed construct normally produces one actionable diagnostic,
  without redundant `VB014` output.
- Independent defects on different lines or colon-separated ranges continue to
  produce their own diagnostics, including generic `VB014` where evidence is
  ambiguous.
- Lint, LSP, and preflight consume the same registry metadata and precedence
  behavior; no grammar modification is needed to force a diagnostic.
- New compile-equivalent syntax rules require VBE rejected evidence and
  accepted negative controls before binding in the fixture contract.

## Alternatives Considered

1. **Suppress all `VB014` when any specific rule fires.** Rejected because it
   hides unrelated malformed syntax and violates the fallback contract.
2. **Use only line-based suppression.** Rejected because colon-separated
   statements can contain independent defects on one line.
3. **Change the grammar to add named error nodes.** Rejected when existing
   recovery ranges already provide reliable evidence; ambiguous shapes remain
   generic by design.
4. **Create one umbrella syntax code.** Rejected because branch, Select/Case,
   Open, and TypeOf have different remediation and evidence contracts.

## Evidence

- `internal/lint/parser_recovery_precedence_test.go` covers unrelated and
  overlapping ranges, including colon-separated statements and `VB059`.
- `internal/lint/conditional_branch_diagnostics_test.go`,
  `select_case_diagnostics_test.go`, and
  `open_typeof_syntax_diagnostics_test.go` cover minimal, valid, nested, and
  ambiguous forms.
- VBE oracle cases `issue594-conditional-*`, `issue594-select-case-*`,
  `issue594-open-*`, and `issue594-typeof-*` bind `VB062`-`VB065` with
  accepted controls and confirmed cleanup.
- ADR-0021, ADR-0024, ADR-0026, and ADR-0032 define the shared IR, registry,
  VBE authority, and corpus-evidence boundaries.

## Related

- Issue #594
- `docs/specs/cli-contract.md`
- `docs/specs/vbe-oracle.md`
- ADR-0021, ADR-0024, ADR-0026, ADR-0032
