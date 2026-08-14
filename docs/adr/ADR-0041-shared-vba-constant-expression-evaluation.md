# ADR-0041: Shared VBA Constant-Expression Evaluation

## Status

Accepted

## Context

Declaration validation needs to answer whether a VBA expression is a value that
can be known without running a procedure. Optional parameter defaults, Enum
members, array bounds, `ReDim` bounds, and literal arithmetic previously used
separate integer or text heuristics. Those heuristics disagreed about suffixes,
operator precedence, forward Const references, and whether an unresolved value
was an error or merely not provable. Batch analysis and LSP also built
name-only project constant tables, so the same source could receive different
results depending on the entry point.

The evaluator is static-analysis infrastructure. It must not load Excel, call
COM, or execute arbitrary VBA while deciding a diagnostic.

## Decision

Add `internal/vba/constexpr` as the shared, presentation-independent evaluator.
It exposes a typed scalar `Value`, a case-insensitive immutable `Environment`,
and `Result` with the three independent outcomes `Known`, `Unknown`, and
`Invalid`. `Known` carries the original typed value and retains the historical
integer projection used by existing bound analyses.

The parser implements a deliberately conservative VBA subset: numeric and
string/date/Boolean literals, intrinsic literal values, resolved Const and Enum
symbols, parentheses, unary operators, arithmetic and Boolean operators,
comparisons, and string concatenation. Calls, runtime-dependent names,
unsupported comparison semantics, unresolved symbols, conditional uncertainty,
and safely unmodelled overflow return `Unknown`. Malformed syntax and domains
such as division by zero return `Invalid`. The evaluator never invokes a call
expression.

Project environments are immutable snapshots assembled from the same parsed
source used by batch and LSP analysis. Public/Friend module constants, Enum
values (including implicit and forward references), and static TypeLib values
are projected into qualified and, only when unique, unqualified names.
Ambiguous or incomplete names remain available to name-only resolver logic but
are filtered from value evaluation. Local source values are overlaid for the
current document without mutating the project snapshot.

Existing `Optional`-default, fixed-array, and `ReDim` bound consumers call the
shared API through adapters. Their diagnostic ownership and codes do not
change; unknown values continue to fail open.

## Consequences

- Const, Enum, optional-default, declaration-bound, and ReDim checks share one
  deterministic grammar and result model.
- Batch and LSP can evaluate the same revision from the same value-bearing
  snapshot, while incomplete snapshots remain conservative.
- Typed strings, Booleans, floating-point values, currency, suffixes, and
  checked integer boundaries are represented explicitly instead of being
  coerced into `int`.
- The supported subset intentionally misses runtime functions, `Like`/`Is`
  semantics, conditional-compilation execution, and other context-dependent
  VBA behavior. Those cases are `Unknown`, not false compile errors.
- Source projection remains a temporary compatibility adapter until
  declaration initializers and Enum value expressions are first-class IR facts;
  the evaluator API itself does not retain parser nodes or source buffers.

## Alternatives Considered

1. **Execute VBA or ask Excel for a value.** Rejected because analysis must be
   deterministic, snapshot-safe, Excel-free, and safe in CI/LSP processes.
2. **Keep independent integer parsers in each rule.** Rejected because it
   duplicates precedence, overflow, suffix, and ambiguity policy and caused
   batch/LSP drift.
3. **Treat every unresolved expression as invalid.** Rejected because an
   unresolved, runtime-dependent, or conditional value is not evidence of
   malformed source; collapsing it would create false positives.
4. **Model every VBA coercion and intrinsic call immediately.** Rejected for
   the first increment; unsafe or context-dependent semantics are explicitly
   `Unknown` until they can be specified and tested.

## Evidence

- Issue #595 acceptance criteria and boundary-value corpus in
  `internal/vba/constexpr/constexpr_test.go` and
  `internal/lint/constant_values_test.go`.
- Existing declaration and array contracts in
  `internal/lint/signature_diagnostics_test.go`,
  `internal/lint/array_shape_diagnostics_test.go`, and
  `internal/analyze/analyzer_test.go`.
- `internal/vba/intel` and `internal/lspserver` project-snapshot tests verify
  that shared lint and compile-equivalent paths retain their existing
  batch/LSP convergence.

## Related

- Issue #595
- ADR-0021 (procedure analysis IR)
- ADR-0040 (shared array-operation shape validation)
- `docs/specs/vba-constant-expression-evaluation.md`
- `docs/specs/vba-analysis-ir.md`
