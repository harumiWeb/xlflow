# VBA Control-Flow Graph

This specification defines xlflow's conservative, protocol-neutral VBA
control-flow graph (CFG). ADR-0022 records the rationale. The graph is an
internal analysis contract built on the procedure IR from ADR-0021; it does not
change public CLI, configuration, JSON, diagnostic, or LSP contracts.

## Construction and Ownership

`internal/vba/cfg.BuildDocument` builds procedure graphs from a
`procedureir.DocumentIR`. It consumes only Go-owned normalized values and never
reads or retains tree-sitter nodes, parser trees, source buffers, or LSP
protocol values.

Each procedure is built independently. A graph and the document-level result
support defensive cloning: callers may mutate returned slices without changing
the builder result or a snapshot's cached copy. Graph construction is
deterministic for the same IR.

### Validation facts

ProcedureIR retains the source ranges and normalized operands needed by CFG
validation: label definitions and transfer targets, `For`/`For Each` control
variables, `Next` variable lists, and statement conditional-compilation and
parser-recovery markers. Clone, rebase, and incremental artifact reuse preserve
these values.

The CFG records protocol-neutral `ValidationFact` values for certain
procedure-local legality failures. A fact identifies its kind, statement ID,
source range, expected/actual values where applicable, and whether the result
is certain. Label resolution is case-insensitive and distinguishes unique,
undefined, duplicate, and recovered targets. The active structured-block stack
proves `Next`, loop `Exit`, and enclosing procedure `Exit` compatibility.
Facts contain no diagnostic code or display text; `ValidationDiagnostics` is
the shared projector for `VB055` (duplicate label), `VB056` (undefined label),
`VB057` (mismatched `Next` variable), and `VB058` (invalid `Exit`).

Recovery and conditional compilation are fail-open: if target resolution or
nesting cannot be proven, no semantic validation fact is projected and the
existing parser-recovery or unknown-flow behavior remains authoritative.

Graph and block IDs belong to one document revision. They are assigned in
stable source order, but are not persistent identities and must not be reused
as workspace-global references after an edit.

## Graph Model

A procedure graph contains source blocks and these synthetic blocks:

- `Entry`, the only procedure entry;
- `NormalExit`, for ordinary fallthrough and `Exit Sub`, `Exit Function`, or
  `Exit Property`;
- `ExceptionalExit`, for errors that leave the procedure;
- `TerminationExit`, for standalone VBA `End`; and
- `UnknownExit`, for unresolved control flow that may leave the known graph.

Every source statement belongs to a block or is represented by the transition
whose control marker it defines. Blocks retain source-order statement
references and canonical `ast.Range` values from the procedure IR. Reachable
and unreachable blocks remain in the graph; construction does not discard dead
source.

Edges record:

- a source and destination block;
- the control transition kind, such as sequential, branch, loop, jump, exit,
  resume, or fault;
- a flow class of normal or exceptional;
- an independent `uncertain` flag;
- the originating statement ID, when a source statement caused the edge; and
- the originating `ast.Range`.

Flow class and certainty are orthogonal. An exceptional edge is not inherently
uncertain, and an unresolved normal jump is both normal and uncertain.

## Structured Control Flow

Sequential executable statements flow in source order unless a terminating
transition replaces that successor.

Multiline and single-line `If` constructs create true and false paths.
Single-line `then` and `else` statements use their normalized branch membership,
not physical-line heuristics. `ElseIf` contributes another conditional branch,
and every branch rejoins at the statement following the complete construct.

`Select Case` branches from the selector to each case. `Case Else` is the
default alternative. When no `Case Else` exists, a no-match path reaches the
statement after `End Select`. Case bodies rejoin after the selection unless
they terminate elsewhere.

`For`, `For Each`, `While`, and all pre-test and post-test `Do While` /
`Do Until` forms include loop-entry, loop-back, and loop-exit transitions.
Pre-test loops include a zero-iteration path. Post-test loops execute their body
before the condition. `Exit For` and `Exit Do` target the successor of the
nearest enclosing loop of the matching kind; an unresolved enclosing loop is
represented conservatively.

`Exit Sub`, `Exit Function`, and `Exit Property` target `NormalExit` when they
match the current procedure kind. Ordinary procedure fallthrough also targets
`NormalExit`. Standalone `End` targets `TerminationExit` and has no sequential
successor.

## Labels, Jumps, and Recovery

Named and numeric line labels are distinct statements from executable source on
the same physical line. Label matching is case-insensitive. A uniquely resolved
`GoTo` targets the corresponding label block, whether the target is forward or
backward.

A missing, duplicate, or recovered target is not resolved to an arbitrary
candidate. Its transition enters conservative unknown flow, preserving the
possibility of reaching relevant known successors and `UnknownExit`. Unknown or
recovered control syntax follows the same rule whenever omitting it could make
a guarantee appear true. The graph records these sources once instead of
materializing an edge from every unknown source to every statement. Queries
interpret a reachable unknown-flow source conservatively. Valid executable
statements whose control behavior is known to be ordinary fallthrough do not
become unknown flow merely because their statement kind is otherwise
unclassified.

Colon-separated executable statements remain individually ordered statements.
Control transitions use their exact normalized ranges so CRLF, UTF-8 source,
and multiple statements on one line do not change graph meaning or diagnostic
locations.

## Error Modes and Exceptional Flow

The builder propagates the active error mode forward and separately along each
path:

- default or `On Error GoTo 0`: an error leaves through `ExceptionalExit`;
- `On Error GoTo <label>`: an error transfers to the uniquely resolved handler,
  or to conservative unknown flow when the label cannot be resolved safely;
  and
- `On Error Resume Next`: an error continues at the successor of the faulting
  statement.

Merging paths preserves every possible active mode rather than selecting one.
`On Error` statements change the mode only on paths that execute them.

Every executable statement remains a possible fault site. The effect-summary
layer in `docs/specs/vba-effect-analysis.md` does not narrow this conservative
fallback. The graph adds an uncertain
exceptional transition according to every error mode that can reach that
statement. Labels and purely structural markers are not fault sites.

`Resume <label>` targets a uniquely resolved label. Bare `Resume` and
`Resume Next` have dynamic destinations determined by the active error and
fault site; when one exact destination cannot be proven, they use conservative
unknown flow. Missing, duplicate, or recovered resume labels are handled like
unresolved `GoTo` targets. A uniquely resolved `Resume <label>` carries the
enabled handler mode to the continuation, where the handler is inactive and may
be entered again by a later fault.

## Query Semantics

The CFG provides procedure-level queries for:

- reachable and unreachable blocks;
- transitions to each synthetic exit;
- block dominance;
- definite assignment at block boundaries; and
- whether one of a supplied set of cleanup statements occurs on every path to
  the selected exits.

Guarantee-oriented queries always include uncertain edges in the selected flow
view. Definite assignment uses intersection at path joins, and an assignment is
definite only when it occurs on every included path. Dominance and all-paths
cleanup use the same conservative reachability. A path through unknown flow or
`UnknownExit` prevents a guarantee unless the queried property is already
established before that path diverges.

The default guarantee view includes normal and exceptional flow. A consumer may
explicitly request a narrower view when its rule defines one. Narrowing by flow
class does not remove uncertain edges of the retained class.

## `VBA204` Integration

`VBA204` error-handler fallthrough detection uses an explicit normal-flow view
that excludes exceptional edges. A handler is a fallthrough risk when its label
block has a reachable implicit normal-flow predecessor. An explicit
`GoTo <handler>` does not count as fallthrough, and a transfer caused only by an
error does not make the handler normally reachable. In particular, an
`Err.Raise` block is not an implicit normal predecessor of the immediately
following handler label; alternate conditional-compilation paths remain
eligible normal predecessors.

The cleanup-label exception is case-insensitive and accepts the existing exact
names (`Cleanup`, `clean_up`, `Finally`, and `Done`) plus qualified labels whose
final underscore-delimited component is `Cleanup` or `clean_up`, such as
`auth_Cleanup` and `AutoProxy_Cleanup`.

This replaces the previous preceding-text heuristic and correctly accounts for
structured branches and jumps. The existing cleanup-label exception, inline
suppression, diagnostic position, message, reason, suggestion, severity,
ordering, CLI JSON, and LSP UTF-16 range remain compatibility requirements.

## `VBA210` Integration

`VBA210` consumes the procedure CFG for `Function` and `Property Get`
procedures. It checks every reachable path to `NormalExit`, including explicit
`Exit Function` / `Exit Property`, structured branches, error-handler paths that
return normally, and shared cleanup or finalization labels. Exceptional,
termination, and unknown exits are not themselves return exits, but exceptional
edges remain in the conservative query so an error handler that later reaches
`NormalExit` is included.

The implicit procedure-name return slot is a local variable in ProcedureIR.
Default VBA initialization does not satisfy the rule. An object return requires
`Set <ProcedureName> = ...`; a known value return requires ordinary assignment
or `Let`. `Variant`, `Any`, and unresolved return types do not receive a
syntactic type assertion. Invalid assignment syntax is not counted as a valid
return assignment; the existing `VBA103` and `VB037` diagnostics remain the
syntax-specific diagnostics for the batch and source-intelligence surfaces.
Known non-returning `Err.Raise` statements are followed only through their
exceptional edge; their normal textual continuation is not a return path.

The non-returning classification is a shared CFG query rather than a
`VBA210`-specific text check. `Err.Raise` and the VBA `Error` statement have no
normal continuation in consumers that classify error outcomes. `VBA237` uses
the same query when distinguishing an explicit rethrow from a handler or
cleanup path that returns normally. Handler `Resume Next` remains a resume
transition with a dynamic destination and is distinct from enabling the
`On Error Resume Next` error mode.

The analyzer uses conservative definite assignment, so an assignment that
dominates the normal exit or occurs in a shared cleanup block satisfies every
path. When a guarantee cannot be established, `VBA210` retains one
procedure-local finding and includes a deterministic representative exit
witness where possible. The rule remains opt-in, batch-only, and does not add a
new public diagnostic ID.

## Snapshot Cache

An `internal/vba/intel.AnalysisSnapshot` owns one lazy document CFG cache for
its immutable revision. The cache:

- builds from the snapshot's cached procedure IR;
- performs at most one build and caches either its result or error;
- is safe for concurrent readers;
- returns defensive copies; and
- retires with the snapshot without retaining parser state.

An incremental successor may reuse completed procedure graph fragments when the
procedure source hash and module-context hash both match. Cached ranges are
rebased from the procedure declaration to its current location; graph, block,
and edge IDs remain procedure-local and deterministic. Module declaration or
conditional-compilation changes, parser recovery ambiguity, and close/reopen
invalidate reuse. Canceled or failed graph builds remain retryable and are not
published into the inherited store.

## Compatibility and Layer Boundaries

The procedure IR owns normalized syntax, statement identity, ranges, and
recovery metadata. The CFG owns executable-path structure and procedure-local
path queries. `internal/vba/effects` owns direct and transitive effect summaries
without changing this graph's conservative potential-fault-site fallback.

Validation-fact ownership follows the same boundary: IR owns operands and
source evidence, CFG owns resolution and certainty, and lint/analyze/LSP
adapters only project the shared facts. The adapters must not reparse control
flow or infer legality from diagnostic text.

The CFG does not define interprocedural call propagation, public visualization,
new CLI output, configuration, or LSP capabilities. Adapters remain responsible
for converting `ast.Range` byte columns to LSP UTF-16 positions.
