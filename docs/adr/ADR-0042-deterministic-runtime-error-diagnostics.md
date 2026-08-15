# ADR-0042: Deterministic VBA Runtime-Error Diagnostics

## Status

Accepted

## Context

Issue #596 adds a class of analysis findings that is different from both a
VBE compile rejection and a general runtime-safety warning. A VBA expression
can be valid source and compile successfully, while its value and the facts
at the use site prove that execution will fail. Examples include dividing by
a known zero, indexing a known array bound outside its range, and reading a
dynamic array that is unallocated on every path reaching the access.

Treating those findings as ordinary warnings loses the distinction between a
possible failure and a failure that is already determined. Treating them as
compile-equivalent errors is also incorrect: source preflight must not claim
that Excel/VBE will reject code that actually compiles. A third category is
needed in the shared rule registry so CLI, LSP, configuration, suppression,
and documentation consumers retain the same meaning.

The shared constant-expression evaluator from ADR-0041, the conservative CFG
from ADR-0022, the procedure-local dataflow boundary from ADR-0025, and the
array state from ADR-0028/ADR-0040 already provide the facts needed for this
decision. Reimplementing local walkers in each diagnostic would make batch and
real-time analysis disagree and would turn unknown `Variant`, late-bound, or
external values into false guarantees.

## Decision

### Evidence class

Amend the registry contract from ADR-0024 with a `runtime-error` evidence
class. Its policy is:

- `error` severity, because the reported execution failure is deterministic;
- `compile_equivalent = false`, because the source may compile in VBE;
- `preflight_blocking = false`, so `push`, `run`, and other workbook commands
  are not stopped solely by this finding;
- inline suppression enabled, subject to the normal analyzer suppression
  policy; and
- default-enabled only when the rule's precision contract is high enough to
  prove the runtime failure.

`compile-equivalent` remains the class for VBE-rejected source and retains its
unsuppressible, preflight-blocking error contract. `runtime-safety` remains the
warning-level class for a failure that is possible or insufficiently proven.
The evidence class, rather than the word `error` alone, determines whether a
finding blocks source preflight or may be suppressed.

### Rule ownership

Add `VBA249` as the umbrella analyzer rule for deterministic runtime errors.
Its canonical configuration key is
`detect_deterministic_runtime_errors`. It is procedure-local, available to
batch and real-time analysis, and emits the additive `runtime_error` context.
The stable `runtime_error.kind` values for the first increment are:

- `division_by_zero`;
- `numeric_type_mismatch`;
- `conversion_type_mismatch`;
- `conversion_overflow`;
- `array_subscript_out_of_bounds`; and
- `array_unallocated`.

`VBA249` owns the deterministic projection only. Existing compile-equivalent
rules, such as `VBA228` and `VBA229`, retain their source-preflight contracts;
`VBA227` retains possible/unknown array lifecycle warnings. An operation must
not produce both a `VBA249` finding and a duplicate lower-confidence finding
for the same runtime fault site.

### Shared facts and precision

The analyzer evaluates a procedure after the shared CFG/dataflow facts reach a
fixed point. It reuses `constexpr.Result`, resolved type information, branch
reachability, definite assignment, allocation state, and known array bounds.
No runtime-error rule may introduce a second procedure-local CFG or execute
VBA/Excel.

Only a value known on every relevant reachable path may support a finding. A
join that contains different values, an uncertain edge, an unresolved or
external call, a possible `ByRef` mutation, an unknown `Variant`, or a
late-bound object invalidates the proof and remains silent. Unreachable code
is not a relevant path. Exceptional and recovered flow keeps the pre-statement
fact unless the shared analysis proves a stronger state.

The first increment covers known zero divisors, known nonnumeric values used
by numeric operators, range-checked explicit conversions, known array bounds,
and dynamic-array allocation state. API-specific argument domains,
object-only operators, general `TypeOf ... Is` impossibility, incompatible
array/scalar binary operators, and late-bound dispatch remain follow-up
contracts until their type/runtime surfaces can be specified with the same
precision.

### Surface and evidence behavior

`VBA249` findings keep the normal analyzer finding envelope and add only the
rule-specific `runtime_error.kind` context. `analyze` and LSP expose the
`error` severity; an analyzer result still has the normal finding exit status.
Source preflight ignores the finding for its blocking decision. The CLI and
LSP must continue to distinguish `runtime-error` metadata from
`compile-equivalent` metadata even when both are rendered with `error`
severity.

Runtime-only findings do not require a VBE oracle case: the oracle answers
whether source compiles, while this contract starts from source that can
compile and proves a later execution failure. If an implementation discovers
that a proposed check is actually a VBE compile rejection, it must move the
finding to the appropriate compile-equivalent rule and bind accepted/rejected
oracle controls before changing its contract.

## Consequences

- Deterministic runtime failures are visible as high-signal errors without
  turning valid VBA into a preflight-blocking compile claim.
- Batch, LSP, and future diagnostic surfaces share one constant/fact model and
  one conservative unknown policy.
- `VBA249` can fail an explicit `analyze`/`check` invocation while workbook
  automation remains available for projects that intentionally suppress or
  review the finding.
- The rule may miss failures that depend on runtime values, external APIs,
  locale-dependent conversion, or unresolved dispatch. This is intentional;
  certainty is preferred over speculative coverage.
- The registry schema, generated diagnostic catalog, analyzer documentation,
  and suppression consumers must recognize `runtime-error` as a distinct
  evidence class. Existing compile-equivalent consumers must not infer
  preflight behavior from severity alone.

## Alternatives Considered

1. **Keep deterministic failures under `runtime-safety` warnings.** Rejected
   because consumers could not distinguish a proven failure from a possible
   one, and users would lose an actionable error signal.
2. **Classify every static runtime failure as compile-equivalent.** Rejected
   because valid VBA can compile and fail only when the expression executes;
   blocking source preflight would misstate VBE behavior.
3. **Report any path on which a failure is possible.** Rejected because the
   issue requires deterministic or strongly proven findings and this policy
   would warn on unknown `Variant`, late-bound, and branch-dependent values.
4. **Build a separate dataflow pass inside `VBA249`.** Rejected because CFG,
   constant, allocation, and type facts would drift from existing rules and
   batch/LSP results would diverge.
5. **Use Excel/VBE during ordinary analysis.** Rejected by ADR-0026 and the
   source-only analyzer boundary; production diagnostics must remain
   deterministic, parallel-safe, and Excel-free.

## Evidence

- Issue #596 defines the candidate runtime failures, precision policy, and
  compile/runtime metadata distinction.
- `docs/adr/ADR-0041-shared-vba-constant-expression-evaluation.md` and
  `docs/specs/vba-constant-expression-evaluation.md` define the shared value
  model and fail-open `Unknown` behavior.
- `docs/adr/ADR-0022-conservative-vba-control-flow-graph.md`,
  `docs/adr/ADR-0025-conservative-vba-source-sink-dataflow.md`, and
  `docs/adr/ADR-0028-array-lifecycle-and-dimension-safety.md` define the
  reusable CFG, dataflow, and allocation boundaries.
- `docs/adr/ADR-0024-shared-static-analysis-rule-registry.md` defines the
  canonical rule metadata and generated CLI/LSP/documentation projections.
- `docs/specs/vba-runtime-error-diagnostics.md` records the public `VBA249`
  contract and positive/adversarial fixture requirements.

## Amends

- ADR-0024 (evidence-class severity and preflight policy)

## Related

- Issue #596
- ADR-0013 (analyzer ownership of semantic runtime-risk checks)
- ADR-0022, ADR-0025, ADR-0028, ADR-0040, ADR-0041
- `docs/specs/cli-contract.md`
