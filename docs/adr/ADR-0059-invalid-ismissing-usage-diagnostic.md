# ADR-0059: High-Precision IsMissing Usage Diagnostic

## Status

Accepted

## Context

Issue #818 (sub-issue of #816) asks xlflow to identify legal but ineffective
`IsMissing` calls. VBA uses this function to detect an omitted `Optional
Variant` argument of the containing procedure. Passing a required or typed
parameter, local variable, field, or computed expression cannot establish
that omitted-argument state. Optional Variant parameters retain the omitted
argument state even when they declare an explicit default; typed optional
parameters can use their default value but do not carry the Variant missing
flag.

The call resolver already distinguishes built-in-like calls from project
procedures, non-callable declarations, and unknown calls. Procedure IR retains
argument expressions, parentheses, source ranges, and recovery metadata, so
the rule can be precise without text scanning or caller analysis.

## Decision

Add one opt-in, procedure-local warning, `VBA283`, configurable with
`detect_invalid_ismissing_usage`. It is available in batch, realtime, and LSP
analysis, supports inline suppression, and does not block source preflight.
The rule registry owns its metadata; the diagnostic remains disabled by
default while corpus evidence accumulates.

Treat a call as the VBA intrinsic only when it resolves as builtin-like and
is either unqualified or explicitly qualified with `VBA.`. Do not diagnose a
same-named project procedure, non-callable shadow, arbitrary receiver, or
unresolved call. A valid argument must reduce through parentheses to a direct
identifier naming a current-procedure parameter that is Optional Variant and
is neither an array nor a ParamArray. Explicit defaults remain valid for
Optional Variant parameters because `IsMissing` observes whether the caller
passed the argument, not whether its effective value equals the declared
default.

Treat parentheses as transparent around a parameter argument. The runtime
behavior of forced parenthesized evaluation was not confirmed in the local
Excel probe, so this is a precision-first analyzer boundary and not a claim
that every parenthesized form preserves the missing-state bit. Report member
expressions and other non-parameter expressions; fail open on malformed
arity, recovered or conditional syntax, and incomplete resolution.

Use a dedicated procedure feature and projection so procedures without a
candidate `IsMissing` call do not run this rule. The implementation is a
single scan over the current procedure's call facts and does not add
dataflow, CFG, or interprocedural state.

## Consequences

- The rule reports the argument expression and explains the failing
  eligibility condition, while callers can opt in independently of other
  diagnostics.
- Resolver-based intrinsic identity prevents project shadowing and arbitrary
  receiver calls from being reported as VBA behavior.
- Parentheses may conservatively hide some ineffective calls until runtime
  evidence is available; accepting them avoids an unsupported false-positive
  claim.
- Corpus workspaces enable the rule for evidence collection without changing
  the production default.

## Alternatives Considered

1. **Match `IsMissing` text in source.** Rejected because comments, shadowed
   procedures, non-callable names, and arbitrary members can share that text.
2. **Treat any Variant expression as valid.** Rejected because `IsMissing`
   observes an omitted optional argument, not whether an arbitrary Variant
   contains a value.
3. **Warn on every parenthesized argument.** Rejected because the local
   runtime behavior was unavailable and the issue requires avoiding
   unsupported false positives.
4. **Enable by default.** Rejected until reviewed corpus evidence supports
   that rollout.

## Evidence

- [Issue #818](https://github.com/harumiWeb/xlflow/issues/818) defines the
  argument eligibility and acceptance cases.
- The [Microsoft `IsMissing` reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/ismissing-function)
  documents its use for optional Variant procedure arguments.
- The [Microsoft named and optional arguments example](https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/understanding-named-arguments-and-optional-arguments)
  declares an Optional Variant with a default and checks that argument with
  `IsMissing`.
- `internal/analyze/ismissing.go` consumes Procedure IR and resolver facts;
  `internal/analyze/ismissing_test.go` covers eligibility, resolution,
  suppression, and surface parity.
- `internal/analyze/procedure_planner.go` gates the rule by a dedicated call
  feature and projection; `internal/staticanalysis/rules/registry.json`
  defines the shared rule contract.
- `docs/specs/vba-ismissing-diagnostics.md` records the user-facing semantic
  and configuration contract.

## Related

- Issue #816 (parent)
- Issue #818
- ADR-0024, ADR-0032, ADR-0046
- `docs/specs/vba-ismissing-diagnostics.md`
