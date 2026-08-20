# VBA Error Suppression and Propagation Analysis

This specification defines the direct and interprocedural error-outcome model
used by `VBA237`. ADR-0033 records the architectural rationale and ADR-0044
defines the bounded fixed-point implementation boundary. The analysis is
protocol-neutral and consumes resolved procedure IR, conservative CFGs, and
the deterministic procedure-effect call graph.

## Public Diagnostic Contract

`VBA237` reports a procedure or call boundary that loses failure information.
It is a default-enabled, high-precision, interprocedural `warning` available to
`xlflow analyze` and the LSP Full diagnostic pass. It is non-blocking for
source preflight, has no automatic fix, and may be suppressed with
`xlflow:disable-line VBA237`, `xlflow:disable-next-line VBA237`,
`[analyze].disabled_rules = ["VBA237"]`, or the compatibility configuration
key `detect_error_suppression_propagation = false`.

The rule does not add fields to analyzer JSON or LSP diagnostics. It uses the
existing finding location, message, reason, and suggestion fields. Fast LSP
diagnostics never compute, cache, rebase, or publish `VBA237`.

## Error Outcome Summary

Each `effects.ProcedureSummary` contains a provenance-bearing `ErrorSummary`.
The public names below describe derived summary properties; the implementation
retains the evidence and path outcomes needed to justify them:

- `has_error_handler`: a reachable `On Error GoTo <label>` has a uniquely
  resolved handler target;
- `uses_resume_next`: a reachable `On Error Resume Next` mode exists; a handler
  `Resume Next` is recorded separately as explicit recovery;
- `suppresses_errors`: at least one reachable error path returns normally
  without rethrowing, explicit recovery, an intentional fallback value, or a
  failure result that callers can observe;
- `rethrows_errors`: at least one relevant handled-error path explicitly exits
  through `Err.Raise` or the VBA `Error` statement; a simultaneous
  `suppresses_errors` outcome records any other path that returns silently;
- `returns_success_flag`: a Boolean `Function` or `Property Get` returns
  explicit success and failure values on the corresponding normal paths;
- `may_raise`: a reachable `Err.Raise` or `Error` statement, an unhandled CFG
  path to `ExceptionalExit`, or a propagated uniquely resolved callee outcome
  may leave the procedure exceptionally;
- `logs_and_continues`: error information reaches a recognized logging sink
  and the path then returns normally without rethrowing or returning failure.

Properties are not inferred from spelling or statement presence alone. For
example, one `Err.Raise` does not make `rethrows_errors` true if another
handled-error path falls through, and a procedure named `TryOpen` is not a
success-return helper unless its Boolean outcomes satisfy the CFG contract.

The fixed-point implementation stores the finite outcome membership separately
from the detailed evidence. Evidence identifies the originating procedure,
canonical source range, statement or call identity, outcome kind, and relevant
handler, cleanup, or return slot. A minimal deterministic witness seed is kept
for every outcome that can be projected into a diagnostic; transitive error
evidence and call-chain prefixes are reconstructed lazily. Evidence identity is
stable and finite so fixed-point propagation does not inflate recursive or
diamond-shaped paths.

## CFG Outcome Classification

For each reachable error transfer, follow exceptional flow into its handler or
resume-next continuation and classify reachable terminal outcomes:

- `rethrow`: `Err.Raise` or `Error` reaches `ExceptionalExit` without a normal
  textual successor;
- `failure_return`: a Boolean return slot is explicitly assigned the failure
  value before `NormalExit`;
- `fallback_return`: a Function or `Property Get` explicitly assigns its return
  slot (including a UDT return-slot field) on every handled-error path before
  `NormalExit`; a literal fallback that dominates handler activation is also
  an explicit fallback;
- `recovered`: a `Resume`, `Resume Next`, or `Resume <label>` explicitly resumes
  execution;
- `normal_suppressed`: a handled or resume-next error can reach `NormalExit`
  without one of the outcomes above;
- `unknown`: recovered syntax, an unresolved handler/resume target, or unknown
  control flow prevents a high-confidence terminal classification.

`Err.Raise` and `Error` are non-returning on normal flow. Their blocks retain
only the appropriate exceptional transition, matching the `VBA210` CFG
contract. A conditional handler is analyzed per path: explicit rethrow on one
branch does not excuse normal suppression on another.

A cleanup label participates when exceptional flow can enter or pass through
it. Cleanup that reaches `NormalExit` without preserving failure as a rethrow,
failure result, or explicit fallback is suppression of the original error,
even if the cleanup operations themselves succeed. An uncertain cleanup exit
does not prove suppression and remains uncertainty.

## Accepted Intentional Patterns

The following patterns do not produce `VBA237` when all relevant paths satisfy
their stated conditions:

- every handled-error path explicitly rethrows with `Err.Raise` or `Error`;
- a uniquely resolved local wrapper is proven non-returning because every
  normal path rethrows or terminates the program;
- a Boolean Function or `Property Get` explicitly returns `True` on success and
  `False` on handled failure, and its caller observes or propagates the result;
- a Function or `Property Get` deliberately supplies an explicit fallback
  return value on every handled-error path;
- a handler uses a `Resume` form as explicit recovery;
- a handler captures `Err.Number`, handles one explicit nonzero expected code,
  and rethrows the other codes;
- `On Error Resume Next` covers one compatibility probe, its `Err` state or
  probe result is inspected, and error mode is immediately restored with
  `On Error GoTo 0` (an optional `Err.Clear` is accepted). The inspection may
  immediately follow restoration and may pass through a local derived value
  or an untaken sibling branch; or
- a Boolean Function explicitly initializes its fallback and returns the
  single resume-next probe result after restoring error mode; or
- a result is assigned to a Boolean local and that exact value reaches a
  control-flow decision before reassignment or exit.

These acceptances do not make every later operation in a broad resume-next
scope safe. `VBA214` continues to own leaked or overly broad
`On Error Resume Next` scope diagnostics; `VBA237` owns only the boundary where
failure notification is lost.

## Logging Recognition

Recognized direct sinks are `Debug.Print`, VBA `Print #`, and
`XlflowDebug.Log`. A uniquely resolved local wrapper is also a sink when its
summary proves that the same error evidence reaches one of those direct sinks.

The logged value must contain `Err.Number`, `Err.Description`, `Err.Source`,
`Erl`, or a procedure-local value whose direct assignments trace to one of
those expressions. A call is not an error log merely because its name contains
`Log`. Ambiguous, unresolved, external, and dynamic logger calls contribute
uncertainty. Logging does not itself signal failure: if the path then returns
normally without rethrowing or returning failure, it is
`logs_and_continues` and is eligible for `VBA237`.

## Call Propagation and Uncertainty

An outcome propagates only across a reachable call whose resolution is
`matched` with exactly one current project-local procedure. The semantic
outcome membership and uncertainty sets are combined by deterministic set
union to a fixed point; indexed adjacency prevents duplicate edge work. Direct
and propagated evidence remain distinct in the compatibility projection, while
the fixed-point implementation keeps only bounded membership plus the witness
seeds required to materialize that projection. Recursion, mutual recursion,
and diamond graphs must converge without duplicating origin evidence.

Ambiguous, unresolved, external, built-in-unclassified, and dynamically bound
calls do not prove suppression or success. Their existing `CallUncertainty`
provenance is retained in the error summary and may be included as context in
another confirmed finding. Uncertainty alone never emits `VBA237`.

A representative call chain is selected deterministically by stable procedure
identity and source location. When a failure-loss boundary is reachable from a
public procedure or host event, that entry point may be named in the finding's
reason, but no second diagnostic is emitted at the entry point.

Path reconstruction is an explanation step, not part of semantic convergence.
It uses the same stable procedure/edge ordering as the pre-optimization
implementation and must preserve the selected representative chain for equal
input. A missing witness, changed path, or changed uncertainty classification
is a compatibility failure even when the finite outcome membership is the
same.

## Success Result Observation

A Boolean success result from a uniquely resolved local helper is checked when
it is:

- used directly by `If`, `ElseIf`, `Select Case`, or a loop condition;
- assigned to a Boolean local that reaches one of those conditions before any
  reassignment or procedure exit; or
- assigned or returned through the caller's own Boolean return slot.

A standalone call discards the result. Assignment to a Boolean local that is
never used for a decision or return before reassignment or exit also discards
the result. These forms emit `VBA237` at the caller's call expression because
that is where the failure signal is lost.

Passing the result as an argument, storing it in a field or container, or
using it through an unsupported expression is uncertain unless the current
analysis can prove observation. Such uncertainty does not produce a
high-confidence ignored-result diagnostic.

## Diagnostic Ownership and Deduplication

Findings are keyed by loss category, owning procedure, and canonical source
range. Emit at most one finding for that identity:

- handler or cleanup swallowing points to the handler or cleanup label;
- failure loss caused by a resume-next region points to its
  `On Error Resume Next` statement; and
- an ignored Boolean success result points to the caller's call expression.

The reason names the suppressed outcome or ignored callee, the originating
procedure when different, and one representative chain when useful. The
suggestion recommends explicit `Err.Raise`, a failure return checked by the
caller, a deliberate fallback, or narrowing and checking a probe as
appropriate.

`VBA204` continues to own normal fallthrough into an error-handler label.
`VBA210` continues to own missing Function/Property return assignments.
`VBA214` continues to own broad or un-restored resume-next scopes. `VBA218`
continues to own Excel API-specific exceptional and `Variant/Error` contracts.
`VBA237` does not duplicate those findings unless a distinct interprocedural
failure-loss boundary exists.

## LSP Project Summary Cache

The workspace index retains tree-sitter-independent `ProcedureIR` and CFG
values for each complete active entry. It builds one immutable project error
summary and confirmed call-edge set for a workspace revision. A fingerprint
includes direct error outcomes and provenance, call-resolution status,
uncertainty, and confirmed outgoing edges.

Open buffers override filesystem entries. While the newest open-document
overlay is pending, both the saved entry and previous overlay remain masked for
this analysis. The project summary is incomplete, so `VBA237` publication is
deferred until the matching overlay is published. Incomplete, canceled, or
obsolete computations are not stored as successful cache entries.

When a fingerprint changes, compute the transitive caller closure through the
union of the old and new confirmed reverse call graphs. This invalidates
callers for added edges, removed edges, and changes between matched,
ambiguous, unresolved, external, or dynamic resolution. Only affected open VBA
documents are scheduled for another Full diagnostic pass.

Each affected document has a dependency generation. Full analysis captures
that generation and the immutable project-summary revision. If either changes
before publication, the result is discarded and the existing dependent-work
queue schedules a replacement. Unrelated call-graph changes do not invalidate
the document.

Overlay publication, `didClose` restoration of the disk entry,
`workspace/didChangeWatchedFiles`, and file deletion use the same fingerprint
comparison and reverse-dependency invalidation. This supplements, and does not
weaken, the document lifecycle and snapshot-generation guards defined by the
general LSP contract.

## Compatibility Boundaries

The analysis is source-only and does not open Excel, access COM or VBIDE, or
change VBE compilation behavior. It does not change analyzer Finding JSON, LSP
Diagnostic fields, the Fast/Full publication ordering, or the public
procedure-effect summary surface (which remains internal). Effect-summary
worklist and compact-fact counters are developer-facing observability only;
they are emitted through the existing opt-in performance log and do not alter
diagnostic semantics.

Unknown behavior is reported conservatively by withholding a definite
`VBA237`, not by assuming success or treating every handler as incorrect.
