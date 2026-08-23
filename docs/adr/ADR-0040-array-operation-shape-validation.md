# ADR-0040: Shared Array Operation Shape Validation

## Status

Accepted

## Context

ADR-0028 introduced a common allocation and dimension state for `VBA227`, but
the state was originally consumed mainly by indexed access, bound functions,
and `ReDim Preserve`. Issue #593 expands validation to all array-sensitive
operations that can be checked from declarations, procedure signatures, and
known assignments. Independent text checks would disagree on fixed arrays,
dynamic arrays, `Variant` values, and values crossing procedure boundaries.

Some cases are VBA compile-equivalent contracts (for example, assigning to a
`Const` or declaring an impossible constant array bound). Others are valid VBA
source whose failure is only observable at runtime (for example, resizing a
fixed array or iterating a scalar). The analyzer must not turn a runtime-risk
observation into a preflight-blocking compiler claim. Local VBE oracle fixtures
provide evidence for that boundary, while ordinary analysis remains
Excel-free.

Issue #696 identifies the cost of having each array-sensitive diagnostic
reconstruct overlapping allocation, shape, bounds, entry-state, and operation
facts. The reuse boundary must be large enough to amortize that work without
making a diagnostic with intentionally different control-flow conservatism
adopt another rule's state contract.

## Decision

Extend the shared array metadata and transfer model with a reusable shape
classification:

- `scalar`;
- `fixed-array(rank)`;
- `dynamic-array(rank)` or `dynamic-array(unknown)`;
- `Variant`; and
- `unknown`.

Declaration facts, procedure parameters/returns, `Array(...)`, `ParamArray`,
and uniquely resolved project-local array returns feed this classification.
Assignments and `ByRef` calls update it conservatively; ambiguous, external,
mixed, recursive, or unresolved values become `unknown`. A `Variant` remains
unknown unless its array nature and the relevant operation are proven.

Use the same shape information for declaration and constant-bound validation,
`ReDim` and `Erase` targets, `LBound` / `UBound` arguments and dimensions,
`For Each` iterable sources, and assignment/argument shape validation,
including array `ByRef` calls. `VB023` retains ownership of the
`For Each` control-variable contract.

`VBA208` remains the owner of `ReDim Preserve` non-final-dimension safety.
`VBA227` remains the owner of non-blocking array lifecycle, allocation,
dimension, bound, and object-array assignment warnings. Compile-equivalent
diagnostics use canonical registry contracts and error severity; they are not
folded into `VBA227` merely because they mention an array.

The implementation reports only statically provable facts. Dynamic bounds,
unknown array rank, unresolved calls, and uncertain `Variant` state remain
quiet. Fixed arrays are never treated as resizable, and a scalar is never
treated as an array or iterable source. `Erase` transitions fixed arrays to
their reset-but-allocated state and dynamic arrays to unallocated state.

The canonical procedure result is an internal, immutable
`ArrayAnalysisResult`. It is materialized at most once for one procedure
analysis revision after the reusable procedure/module facts have been
prepared. The compact implementation keeps the shared variable/entry-state
preparation and policy-lane outputs in categorized slices
(`lifecycleFindings`, `runtimeFindings`, `redimFindings`, and
`rangeFindings`). Projectors expose copies of those slices, so the result is
still read-only after materialization; the stored `Finding` values do not
carry rule enablement, suppression mutation, or a new severity contract. It
lives only for the current procedure analysis and is safe for concurrent
read-only projection; it is not a project-wide or cross-revision cache.

The array kernel projects the result into the applicable diagnostics for both
batch and realtime/LSP analysis. The block-level projection preserves the
existing `VBA208`, deterministic array `VBA249`, object-array `VBA101` /
`VBA102`, and related operation semantics. The `VBA227` lifecycle projection
retains its physical source-line ordering, normal-edge pruning,
`On Error Resume Next`, and reliable-allocation policy. `VBA241` consumes
shared `ReDim` and loop-region facts without another fixed-point walk. `VBA226`
retains its scalar/2D `Range.Value` shape lattice and may use an explicitly
separate secondary pass when that policy requires it. These policy lanes are
projection choices, not alternative public state models; duplicate findings
continue to be resolved by the existing operation identity and ownership
rules.

## Consequences

- Array-sensitive diagnostics agree across declarations, assignments, calls,
  bounds, `Erase`, `ReDim`, and `For Each`.
- Shared preparation and the main array fixed point run once per procedure
  revision, after which multiple diagnostic projectors read the same result.
- The performance recorder exposes `array_kernel_runs`, `array_cfg_walks`,
  and `array_projection_runs`. Core array projections keep one kernel result
  and one main CFG walk; an applicable `VBA226` secondary pass is explicit.
- Existing `VBA208`, `VBA227`, `VBA228`, and `VB023` ownership boundaries remain
  stable; new compile-equivalent contracts are registered separately from
  runtime-safety warnings.
- Function returns, `Array(...)`, `ParamArray`, and `ByRef` arrays can be
  validated without trusting external or ambiguous code.
- Conservative unknown handling may miss a warning when VBA's runtime value is
  knowable only through execution. This is intentional and avoids false
  positives for `Variant` and dynamic cases.
- The VBE oracle corpus grows only for questions whose compile-equivalent
  meaning is uncertain. Accepted runtime-risk probes remain unbound language
  observations and do not suppress policy or safety diagnostics.
- Real-world corpus snapshots remain observations. A changed diagnostic set
  requires focused review and an explicit, atomic snapshot update; snapshots
  are never edited to hide a regression.

## Alternatives Considered

1. **Add one matcher per operation.** Rejected because operation-local type
   guesses would diverge after assignments, `Erase`, branches, and calls.
2. **Treat every `Variant` as an array or scalar.** Rejected because either
   choice converts uncertainty into a false guarantee.
3. **Make every array misuse compile-equivalent.** Rejected because fixed-array
   `ReDim`, scalar `LBound`, and scalar `For Each` are valid source shapes whose
   failure is runtime behavior.
4. **Use VBE from normal lint/analyze/LSP runs.** Rejected by ADR-0026; the
   oracle remains a local, sequential evidence tool and never a production
   dependency.

## Evidence

- Issue #593 acceptance criteria and test matrix.
- Focused compile-oracle controls and bindings are recorded by case ID:
  `known-compile-accept` / `known-compile-reject` are harness controls;
  `const-assignment-accepted` / `const-assignment-rejected` bind `VB060`;
  and `fixed-array-valid-bound` / `fixed-array-reversed-bound` bind `VB061`.
  `redim-scalar` is a compile rejection and `redim-reversed-bound` is a compile
  acceptance; both remain unbound language observations because VBA227 owns
  runtime safety rather than compile-equivalent diagnostics.
- `docs/specs/array-lifecycle-safety.md` and ADR-0028 for the existing
  allocation/dimension model and `VBA208` ownership.
- Issue #696 acceptance criteria and its shared-kernel performance matrix.
- `docs/specs/vba-analysis-ir.md` and `internal/vba/procedureir` for array
  declaration, bounds, `ParamArray`, and procedure-return facts.
- `docs/specs/vbe-oracle.md` and ADR-0026 for compile-equivalent evidence,
  promotion, negative controls, and infrastructure-failure handling.
- `docs/specs/static-analysis-corpus.md` for snapshot review and explicit
  update contracts.

## Supersedes

- ADR-0028 (array lifecycle state remains valid; this ADR extends its operation
  validation boundary).

## Superseded by

None

## Related

- Issue #593
- Issue #696
- ADR-0026, ADR-0028
- `docs/specs/array-lifecycle-safety.md`
- `docs/specs/range-value-array-shape-safety.md`
- `docs/specs/cli-contract.md`
