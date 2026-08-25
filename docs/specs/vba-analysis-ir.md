# VBA Procedure Analysis IR

This specification defines xlflow's protocol-neutral, procedure-level
intermediate representation (IR). ADR-0021 records the rationale. The IR is an
internal analysis contract: it does not change the public CLI JSON or LSP
diagnostic contracts.

## Construction and Ownership

`internal/vba/procedureir` exposes two construction paths:

```go
func BuildSource(BuildOptions, []byte) (DocumentIR, error)
func BuildParsed(BuildOptions, *ast.ParsedDocument) (DocumentIR, error)
func ResolveView(DocumentIR, Resolver) ResolvedDocumentView
func Resolve(DocumentIR, Resolver) DocumentIR

func (ResolvedDocumentView) ResolvedCall(procedureIndex, callID int) (ResolvedCall, bool)
func (ResolvedDocumentView) ResolvedAccess(procedureIndex, accessID int) (ResolvedAccess, bool)
func (ResolvedDocumentView) ResolvedEvent(procedureIndex, eventID int) (ResolvedEvent, bool)
func (ResolvedDocumentView) Materialize() DocumentIR

func DiagnosticsView(ResolvedDocumentView, bool) []ResolutionDiagnostic
```

`BuildSource` parses the supplied immutable source and closes the parsed
document before returning. `BuildParsed` reads an existing caller-owned parsed
document and never closes it. This is the incremental construction path for
LSP: if the document was produced by `ast.ParseDocumentIncremental`, the builder
uses that already-incremental tree and does not parse the source again.

All tree and source access occurs inside `ast.ParsedDocument.Read`. Returned IR
values own every string and slice they expose and contain no tree-sitter node,
tree, or borrowed `ParsedView.Source` slice. They remain safe to read after the
parsed document is closed.

`DocumentIR` is the immutable syntax-local input for project resolution.
`ResolveView` builds a read-only, revision-scoped overlay containing only
project-dependent call, access, and event resolution facts, including effective
access scope. Its lookup helpers hide whether a fact is stored inline or in a
side table. They do not expose mutable candidate storage; callers that need an
independently owned IR use `Materialize` or the compatibility `Resolve` API.

The overlay is keyed by a deterministic procedure identity composed from path,
module name and kind, qualified procedure name and kind, declaration start
byte, and source-order ordinal. Within that procedure, calls use their stable
call ID and accesses and event references use revision-local IDs. The IDs are
valid only for the document revision that created them and must not be used as
workspace-global or source-text identities. Production IR assigns those IDs;
hand-built IR with missing or duplicate IDs uses the corresponding procedure
slice ordinal as a compatibility fallback. Access and event IDs are internal
metadata omitted from serialized IR/JSON, so adding them does not change the
existing wire shape.

`BuildOptions` supplies document context that syntax alone cannot reliably
provide, including path and module kind. A missing module name is derived using
the same source attribute/path fallback as existing source inspectors.

## Document and Procedure Meaning

`DocumentIR` records:

- path, module name/identity, and module kind;
- exported module attributes such as `VB_Exposed`;
- parse recovery flags corresponding to tree-sitter `HasError` and
  `HasMissing`;
- module-level declarations; and
- procedures in source order.

`ProcedureIR` records a `ProcedureSymbol`, declarations, normalized statements
and expressions, call sites, and variable reads and writes. Its block
information is syntactic parent/child nesting only.

`ProcedureSymbol` distinguishes `Sub`, `Function`, `Property Get`,
`Property Let`, and `Property Set`. It records visibility, parameters, return
type where applicable, declaration and body ranges, and event-handler metadata.
Function and property return slots also retain the normalized value shape
(`scalar`, `Variant`, fixed array, dynamic array, or `unknown`) and parsed array
bounds when the return type declares an array. Module and procedure
declarations retain the same value-shape classification plus `Const`/Enum
constness, so array-sensitive consumers do not need to infer those facts from
rule-specific text scans.
Property accessor kinds stay distinct even when they share the same VBA name.
Parameters are the canonical declaration-signature facts consumed by
declaration validation and projected by symbol/call tooling. Each parameter
retains its effective passing mode and whether it was written explicitly,
Optional/ParamArray modifiers, default-expression text and range, array shape
(`none`, `dynamic`, `bounded`, or `invalid`), parsed array bounds, and recovery
state. Conditional procedure headers carry mutually-exclusive branch identity
so cross-accessor and type compatibility checks compare only declarations that
can be active together. A branch records its source condition for `#If` and
`#ElseIf`; the `#Else` branch has an empty condition and is identified by its
branch number within the shared group.
For a procedure with no body block, the body range is a zero-width range at
the start of its `End Sub`, `End Function`, or `End Property` statement.

A module-level `Event Foo(...)` declaration is a declaration and must not create
a `ProcedureIR`; its parameter list is retained on the declaration. Type and
Enum declarations are also retained as module-level declarations. `Declare Sub` and `Declare Function` are declarations without
VBA bodies; they may participate in symbol or call resolution but are not
procedure bodies.

Event metadata uses module context as well as the procedure name:

- `Auto_Open` and `Auto_Close` are recognized host entry points.
- In `document` modules, `Workbook_*` and `Worksheet_*` procedures are
  recognized document event handlers.
- In `form` modules, a non-test procedure whose name has non-empty text on both
  sides of its last underscore is recognized as a UserForm/control event
  handler.
- Test-shaped names remain tests rather than UserForm event handlers.
- Similar names in ordinary standard modules are not reclassified as document
  or UserForm events.

This classification is syntactic host metadata, not proof that a COM event
source exists or that a handler signature matches a type library.

## Statements, Expressions, and Accesses

Statements have a stable source-order ID, an optional parent ID, a canonical
`ast.Range`, and recovery state. Normalized statement kinds cover:

- declarations and `Let`/implicit or `Set` assignments;
- explicit and implicit calls;
- `If`, `ElseIf`, `Else`, `Select Case`, and `Case`;
- `For`, `For Each`, `Do`, `While`, and `With`;
- `Exit`, labels, `GoTo`, `On Error`, and `Resume`; and
- unknown or recovered syntax.

Nested constructs preserve their source parent/child relationship. Colon-
separated statements are separate statements ordered by byte position.
Multiline logical statements remain one normalized statement whose range spans
the complete source form.

Control-flow statements retain protocol-neutral operands for CFG validation:
label definition and transfer-target ranges, `For`/`For Each` control
variables, and the complete `Next` variable list with one range per supplied
name. Statements also carry conditional-compilation branch identity. Clone and
incremental rebase operations copy/rebase these ranges and values together
with the statement, so downstream CFG facts never need to inspect parser nodes
or reparsed source text.

Expressions normalize identifiers, literals, member access, calls, `New`,
unary and binary operators, parentheses, and unknown syntax. Their relationship
to a statement identifies roles such as assignment target, assigned value,
condition, or argument.

Variable accesses identify their source range, name, access mode
(`read`, `write`, or `read_write`), and resolved scope. A simple value reference
is a read. An assignment target is a write; `Set` remains distinguishable as
object assignment. An operation that both consumes and updates a value uses
`read_write` and must not be reduced to write-only.

The IR is conservative. Unsupported or ambiguous syntax is represented as
unknown/recovered instead of being assigned an invented meaning.

## Scope and Resolution

Symbol scope values are:

- `parameter`
- `local`
- `module`
- `project`
- `unresolved`

VBA name matching is case-insensitive. Within a procedure, parameters and local
declarations shadow module declarations; module declarations shadow project
symbols. Same-scope ambiguity remains explicit rather than depending on map or
filesystem iteration order.

A reference bound to a declaration in the current module has `module` scope
even when that declaration is publicly visible to other modules. Public or
Friend visibility makes the declaration a project candidate for other modules;
it does not allow a project overlay to replace the current-module binding.
For procedure calls without a receiver, cross-module candidates are limited to
standard modules. Procedures in class, document, and UserForm modules require
an explicit receiver outside their own module.

Syntax construction does not require a project snapshot. Calls initially carry
`not_attempted` resolution, and variable accesses that cannot be decided from
procedure/module declarations remain unresolved until a resolver is supplied.

`ResolveView` applies the supplied project snapshot without mutating the base
IR or rescanning source. Replacing the resolver can therefore re-resolve the
same syntax after a workspace symbol change. `Resolve` remains the compatibility
API for consumers that require an independently owned resolved `DocumentIR`;
it materializes equivalent overlay facts and preserves the input-unchanged,
independent-ownership, and nil-resolver behavior of the existing API.

The overlay and materialized forms have identical completeness and ambiguity
semantics. This includes unresolved, incomplete, ambiguous, local/module/
project scope, enum-member ambiguity, dynamic/external calls, event resolution,
and deterministic candidate ordering. Overlay identity never uses source text
when a stable IR identity is available.

Call resolution statuses are:

- `not_attempted`: no project resolver has been applied;
- `matched`: exactly one visible project candidate was selected;
- `ambiguous`: multiple visible candidates remain;
- `unresolved`: no known candidate or special category applies;
- `external`: a known external receiver owns the call;
- `builtin_like`: the call text matches a recognized VBA-like built-in; and
- `member_call`: a receiver/member call cannot be bound conservatively.

Candidates use deterministic order and preserve qualified name, kind, file, and
declaration line needed by the existing `inspect calls` projection. Resolver
symbols also retain module kind so receiver-less calls cannot bind non-standard
module procedures across module boundaries. Private procedures are visible only
from the same module. Resolution does not claim full VBA type inference, COM
type-library binding, or Excel object-model dispatch.

`Diagnostics` and `DiagnosticsView` share one resolution-diagnostic walker.
The walker obtains syntax and recovery state from the base IR and obtains
resolution facts through the view. Batch analyzer `VB052`-`VB054` resolution
diagnostics use `DiagnosticsView` to avoid a second full-document clone. Other
consumers that need independent snapshot ownership may continue to use
materialized `Resolve`; this boundary does not change LSP, effects, lint,
metrics, or public CLI/JSON behavior.

## Ranges and Determinism

Every externally useful syntax fact uses `ast.Range`:

- lines and columns are 1-based;
- columns count UTF-8 bytes, matching tree-sitter positions;
- `StartByte` and `EndByte` are absolute zero-based source byte offsets; and
- ranges use tree-sitter's half-open end position.

LSP adapters convert these byte-based ranges to UTF-16 positions. The IR never
stores LSP positions, and introducing the IR must not change existing CLI or
LSP locations, including CRLF, non-ASCII, and supplementary-plane input.

All slices are deterministic:

- procedures, declarations, statements, expressions, calls, and accesses use
  ascending source byte order;
- ties use stable kind/name keys defined by the builder, never Go map order;
- statement IDs are assigned from the final source order and are stable for the
  same source and build options; and
- resolver candidates use stable qualified-name, kind, file, and location
  ordering.

Resolution overlay entries use the same source-ordered procedure identity and
revision-local call/access/event IDs on every lookup. Candidate slices are
copied only when an owned result is requested, and candidate ordering remains
the resolver's stable qualified-name, kind, file, and location order.

Determinism is per document revision. IDs are not persistent identities across
edits and must not be stored as workspace-global references.

## Recovery

Tree-sitter recovery does not make construction fail by itself. `DocumentIR`
preserves `HasError` and `HasMissing`, emits every complete fact that can be
normalized safely, and represents meaningful unclassified regions with
unknown/recovered statements. Recovered facts retain exact ranges and are
marked so consumers can avoid treating them as certain.

An operational parse/read error still returns an error. A malformed but
successfully parsed document returns partial IR plus recovery flags.

Recovery acceptance belongs to the consumer:

- interactive/editor consumers may use partial IR for best-effort features;
- existing batch `analyze` behavior may reject a document with parser recovery;
  and
- a rule must not turn unknown syntax into a positive control-flow or semantic
  claim.

## Snapshot Cache

An `internal/vba/intel.AnalysisSnapshot` owns the syntax IR for one immutable
document revision. The cache:

- builds lazily from the snapshot-owned `ParsedDocument`, or accepts a
  caller-supplied immutable IR for the exact batch revision;
- is safe for concurrent readers and performs at most one IR build per
  snapshot;
- shares the same parsed document already used by snapshot symbols,
  diagnostics, and calls;
- returns defensive copies so callers cannot mutate cached slices or nested
  candidate data; and
- is retired with the snapshot and never closes the parsed document itself.

Batch analysis may construct the IR before creating the snapshot and seed it
through `intel.NewAnalysisSnapshotWithArtifacts`. The snapshot deep-copies the
IR at that boundary; parsed-document ownership remains unchanged.

An incremental edit creates a new parsed-document snapshot and a new
single-flight document result, but a successor within the same open lifecycle
inherits an immutable store of completed procedure fragments. Fragment identity
includes procedure kind/name/ordinal, exact procedure source, and module
context. Materialization rebases every `ast.Range` from the declaration start,
orders procedures by current source, and reassigns document declaration IDs.
Canceled, panicked, recovered, ambiguous, and in-progress builds are retryable
and are never inherited. A full replacement within the same lifecycle may
reuse fragments; close/reopen may not. After a successful IR or CFG build, the
store retains only the current cache key for each procedure identity. Obsolete
source or module-context variants are pruned so repeated edits cannot grow the
successor store with unreachable revision artifacts.

Module declarations and conditional-compilation directives are hashed
separately from procedure bodies. A module-context change invalidates semantic
IR and CFG reuse. A conditional-compilation change, overlapping/recovered
procedure boundary, or non-unique catalog mapping falls back to a complete
module build.

### Analyzer-derived facts

The analyzer may derive a `moduleAnalysisFacts` value and one
`procedureAnalysisFacts` value for each procedure from a revision-local
immutable view of the canonical `DocumentIR`/`ProcedureIR`. These are
implementation-level projections for one analysis revision, not a second
snapshot cache. A procedure view retains the owned procedure/CFG elements and
analysis overlays; it does not materialize replacement declaration, statement,
expression, call, access, or parameter collections. File facts own compact
declaration, source-ordered constant, procedure-name, and line-ownership
indexes; procedure facts own compact offsets plus read-only ID lookups and
statement-grouped call, access, and member-expression indexes. The file facts
also keep a compact declaration-start-to-procedure-facts pointer index, without
duplicating procedure IR. Empty projections do not eagerly allocate maps.

Facts and views are immutable after construction. Rule families and future
procedure workers may read the same canonical IR through lookup/iterator
helpers concurrently, while flow state, findings, CFG worklists, object
variables, and other rule-specific overlays remain local and mutable. Fact
indexes retain only Go-owned IR references and compact offsets: they must not
retain tree-sitter nodes, `ParsedDocument` instances, source snapshots, or
obsolete revisions. Their lifetime is bounded by the parsed file/procedure
revision and they are released with that analysis result. Snapshot and public
compatibility getters may continue to return defensive copies at their
ownership boundary.

Hot analyzer rules must use the private read-only procedure span and grouped
fact iterators for immutable reads. Procedure iteration uses indexed span
access so large procedure values do not make range-over-function closures
escape. Statement-grouped member-expression iteration preserves canonical
source order and uses the compact grouped index without materializing a
temporary slice. The owned procedure and grouped-member getters remain the
explicit boundary for callers that mutate effect overlays, recover from an
incomplete index, or require an independently owned compatibility value.

Batch and realtime entry points construct and attach facts before rule workers
run. The performance recorder reports `module_fact_builds` and
`procedure_fact_builds`; a normal file revision contributes one module build
and one procedure build per procedure, independent of the number of rule
families that consume them.

Batch resolution diagnostics attach a read-only `ResolvedDocumentView` to the
immutable document facts. The view is procedure-local and contains only
resolution side tables; it does not allocate a second full-size copy of the
procedure IR. A consumer that needs an independently owned resolved snapshot
continues to call `Resolve`/`Materialize`, which remains the compatibility
boundary.

### Module semantic facts

`moduleAnalysisFacts` may also own the immutable module-wide semantic facts
needed by analyzer kernels. They are derived from the canonical
`DocumentIR`/`ProcedureIR` and existing source normalization in one bounded
pass for one module/revision. They are an implementation-level projection,
not a complete normalized AST and not a replacement for the IR or resolution
overlay.

The module option projection includes `Option Private Module` as a tri-state
value: `present` for an exact declaration in complete canonical input,
`absent` when complete input proves no declaration exists, and `unknown` for
recovered, incomplete, or ambiguous option syntax. Comment and string content
cannot establish the option. Consumers must fail open for `unknown`, retaining
the existing possibility that a procedure is externally visible. A module-wide
option check is performed at most once per module/revision; procedure loops
must consume the stored value and must not scan all source lines.

The same immutable pass may index only the operation facts required by
idempotent array setup: canonical lowercase identifier keys and compact
observations for scalar assignment, direct `ReDim`, `Erase`, and whole-array
assignment. This is not an array participant graph or a data-flow state. Shared
maps and slices are private implementation details. Read access uses scalar,
enum, or small owned-copy accessors; no mutable slice or map is exposed to a
kernel.

Module facts are attached before procedure workers run in batch, realtime, and
LSP analysis. They are safe for concurrent readers and immutable after
construction. Their lifetime is the parsed-file/IR revision; edits and
complete/incomplete transitions require rebuilding the facts, and no
cross-revision persistent cache is defined. Compatibility and synthetic
callers must share one caller-owned facts value rather than triggering an
implicit rebuild.

The opt-in performance recorder reports `module_fact_builds` and
`module_option_scans` as stderr-only counters. A normal module revision
contributes at most one of each, regardless of its public procedure count.
These counters are not part of CLI JSON, LSP payloads, or diagnostic output.
`BenchmarkModuleAnalysisFacts` covers small, many-procedure, and
operation-heavy construction/reuse cases; the real-world corpus benchmark
`BenchmarkRealWorldCorpus/ronecone/analyze-only` is used for before/after CPU,
heap, and allocation measurements. Small-module benchmark medians are kept as
a regression guard while giant-module profiles verify normalization reuse.

### Procedure features and applicability planning

Each `procedureAnalysisFacts` value also owns one immutable applicability
summary for the same analysis revision. The summary is a compact internal
value equivalent to two bit sets: `present` contains features proven to occur
in the procedure and `unknown` contains features whose absence cannot be
proved. A feature in neither set is proven absent. The representation and bit
assignments are internal; this specification defines the tri-state contract,
not a public type or serialization format.

Feature classification consumes the already-owned declarations, statements,
expressions, calls, and accesses. It is built as part of the existing facts
pass with at most one bounded visit of each owned IR projection. The planner
then combines the local summary with module/project facts. Neither step
may reparse source or run an independent source-text scan once per
feature or rule family. Features are added only for actual semantic
prerequisites, such as array shape/use and `ReDim`, loops, object/member or
Dictionary/Collection access, error handling, runtime candidates, source/sink
operations (including process launch, SQL, HTTP, and file I/O), resource
acquire/release, Excel access, application-state mutation, event metadata,
calls, ByRef mutation possibility, and dynamic/unresolved-call uncertainty.

The applicability planner maps enabled diagnostic IDs to explicit semantic
domain requirements and combines them with this summary, CFG/resolution
completeness, module facts, and procedure/project effect summaries. It must
run before a domain's expensive setup, CFG walk, or flow-state allocation.
Only a proven-absent requirement may skip a domain. Unknown applicability is
planned. Parse/read errors, missing or recovered IR, unknown CFG flow,
incomplete resolution, ambiguous/dynamic/unresolved calls, and missing
module/type/project facts therefore fail open. A hand-built source procedure
that lacks the owned facts required for classification uses the same
conservative unknown fallback.

When a project capability is required, its builder must preserve the complete
project closure needed by that capability. A domain may use a participant
subset only when the feature seed and resolved caller/callee closure prove that
the reduction preserves semantics; otherwise it retains the complete
procedure set. The project capability planner may skip construction only when
no enabled diagnostic requires the capability. Compile-equivalent diagnostics
and existing unconditional compatibility paths retain their existing behavior.

The summary and planner are reused by batch and realtime/LSP entry points and
are safe for concurrent readers. The performance recorder reports one
`planned_<domain>_runs` or `skipped_<domain>_runs` decision per procedure and
gated domain. These are opt-in stderr telemetry only and do not change normal
CLI JSON, LSP, diagnostic, or configuration contracts.

#### Array participant closure

Array interprocedural planning derives one immutable participant set per
analysis revision after resolution and module facts are complete enough for
the closure. Seeds are procedures with direct array operations or accesses,
array parameters or returns (including `ParamArray`), object-array use, or a
module-array read/write. A module-array declaration with no procedure access
does not seed unrelated procedures.

The closure follows semantically array-shaped resolved callees and reverse
callers, array ByRef arguments/results, array-return chains, shared module-array users, and
initializer/setup/helper edges for class, document, and UserForm modules. It
is built from stable procedure identity and source ordinal; adjacency and the
worklist are sorted before traversal. `inferArrayReturnSummaries`, ByRef and
module allocation, module-entry state, and ByRef-entry state all receive this
same set. A procedure outside the set must not be scanned by an array
interprocedural fixed point when its absence is proven.

For resolved reverse callers, the planner permits one additional hop only when
the target has complete facts and its module contains at most 512 procedures.
This preserves short wrapper/helper chains while preventing a generic call hub
in a giant module from expanding the array closure; a caller with its own
array seed is still included independently. Recovered or incomplete targets
use the bounded uncertainty policy below.

Fail-open is bounded by the best available ownership boundary. An ambiguous
call keeps all resolved candidates and the related SCC; an unresolved or
dynamic call keeps known project-local candidates, or the containing module
when target identity is unavailable. External/member calls preserve unknown
state without adding unrelated project procedures. Recovered/incomplete IR
expands through known procedures, SCCs, and the module-array cluster before
using the containing module. Inputs with no reliable module/file ownership
retain the complete-project fallback for compatibility. This policy preserves
array propagation through module state, ByRef arrays, returns, and
initializers while retaining compile-equivalent `VBA101` and `VBA102` results.

The array fixed point uses a deterministic source-ordinal worklist and keeps
per-participant contributions. Only a changed summary or entry contribution
requeues dependent participants; unchanged participants are not rescanned on
every iteration. Cancellation checkpoints apply to CFG walks; participant
closure construction does not currently accept or check a context.
`walkArrayCFGCombined` retains its existing cancellation behavior while array
fixed-point worklists remain deterministic.

Participant telemetry is developer-only and additive:
`array_participant_procedures` counts unique closure participants,
`array_interprocedural_cfg_walks` counts started CFG walks in array summary or
entry fixed points, and `array_worklist_revisits` counts post-initial
participant evaluations. Existing local `array_cfg_walks` and candidate
counters keep their prior meanings.

### Project semantic capabilities

Procedure applicability is only one input to project analysis planning. After
the owned IR, module facts, resolution completeness, and procedure feature
summaries are available, the analyzer builds an immutable project capability
plan from the enabled diagnostic registry. Each project/domain diagnostic
declares direct capability requirements; the planner computes their transitive
closure before any project-wide semantic builder runs. A rule implementation
must consume the resulting capability bundle and must not construct a missing
project context as a private fallback.

The initial internal capability vocabulary and dependencies are:

| capability             | direct dependencies              |
| ---------------------- | -------------------------------- |
| `TypeDB`               | —                                |
| `Resolution`           | `TypeDB`                         |
| `ProjectConstants`     | `TypeDB`, `Resolution`           |
| `ByRefSymbols`         | `Resolution`                     |
| `Effects`              | `Resolution`                     |
| `ObjectFlow`           | `Resolution`                     |
| `ArrayInterprocedural` | `Resolution`, `ProjectConstants` |
| `DataFlow`             | `TypeDB`                         |
| `DictionaryCollection` | `Resolution`                     |
| `ApplicationState`     | `Effects`                        |
| `EventReentry`         | `Effects`                        |
| `PublicAPITypeIndex`   | `TypeDB`, `Resolution`           |
| `ExcelLoopSymbols`     | `TypeDB`, `Resolution`           |
| `ExcelAPIHelpers`      | `Resolution`                     |
| `ModuleState`          | —                                |

The table is an internal completeness contract, not a public CLI or LSP
schema. A capability is constructed at most once per analysis revision. Batch
analysis may build the planned closure eagerly in dependency order before
procedure workers start. Realtime/LSP analysis may memoize the same immutable
results by revision; concurrent requests for the same revision share them,
while a new revision or a complete/incomplete transition invalidates them.

Capability planning is independent from optional runtime-analysis settings for
compile-equivalent diagnostics. The baseline compiler-equivalent checks and
the unconditional Intel, ByRef, assignment, local-type, CFG, array-shape,
`VB052`--`VB054`, `VBA101`, and `VBA102` paths remain planned when their
existing analyzer contract requires them. Optional effects, object flow,
array interprocedural, data-flow, Dictionary/Collection, application-state,
event-reentry, public-API, and Excel-loop work is built only when an enabled
diagnostic requires the corresponding capability.

Participant filtering is permitted only when the feature seed and resolved
caller/callee closure prove that the reduced set preserves semantics. The
filtered set must include related module state, ByRef state, helper procedures,
class/document/UserForm initializers, and uncertainty boundaries. Recovered IR,
ambiguous or dynamic calls, incomplete resolution, or an unknown module-state
boundary fail open to the complete procedure set. Effects, `ApplicationState`,
and `EventReentry` keep the complete project closure whenever required, while
`PublicAPITypeIndex` may scan only the public API surface.

Capability builders share the already resolved IR and resolver. In particular,
the Effects capability must not rebuild resolution, and a project rule must not
call an independent `buildProjectEffects`-style fallback after planning. The
same capability plan is used by batch and Full realtime/LSP diagnostics so
findings, uncertainty handling, and the compile-equivalent baseline remain
consistent across entry points.

### Procedure execution plans and semantic result dependencies

The main procedure analyzer executes an immutable `procedureAnalysisPlan` for
each procedure and analysis revision. The plan is derived after procedure
features, project capabilities, and effect facts are available. The internal
diagnostic requirement table is the single source of truth for the plan's
kernel, projection, and semantic-result dependency closure. A plan contains
only the applicable kernels and projections; a proven-absent requirement may
skip work, while unknown, recovered, incomplete, ambiguous, or unresolved
inputs remain planned.

Resolved calls use their propagated effect facts plus direct syntax features;
the mere presence of a resolved call is not sufficient to schedule every
domain. Recovered, ambiguous, external, and unresolved call evidence remains
fail-open.

The planner owns applicability, capability requirements, kernel scheduling, and
result dependencies. A semantic kernel owns its CFG walk, fixed point, or
data-flow state and publishes an immutable result. A diagnostic projection owns
diagnostic code, severity, message, reason, suggestion, evidence, and
compatibility filtering. The existing combined source-line scan remains one
shared kernel when its normalized statement observations can serve multiple
textual projections. It must not be split into one source traversal per rule.

Kernel and projection execution follows the executor's static canonical append
sequence. Plan materialization, map iteration, and worker completion cannot
change that order.
Procedure workers write to stable procedure-indexed result slots; the existing
source-order merge, final diagnostic sort, and suppression stages remain the
authority for observable finding order and multiplicity. Batch and Full
realtime/LSP entry points use the same requirement metadata and conservative
ordering semantics.

Semantic results are immutable and scoped to one procedure and analysis
revision. A result is materialized at most once for that plan and can be read
by multiple projections through a procedure-local result store. The store is
released with the procedure analysis or revision; canceled, failed, and
obsolete results are not published or retained. This boundary does not add a
persistent or cross-run cache. Result identity and dependencies may be useful
to a future incremental analyzer, but no incremental or cross-process cache
contract is defined here.

Data-flow execution is split into independent lazy lanes inside that same
procedure-local boundary. The generic lane serves `VBA224`, `VBA236`, and
`VBA239`; the HTTP lane serves `VBA246` and `VBA247`; and the file-path
`VBA245` projection remains independent. The plan records enabled and planned
lane masks separately. A lane with no planned projection does not start its
CFG walk or kernel. Each planned lane is materialized at most once, with HTTP
materialized before generic data-flow so the established `VBA224` ownership
and suppression behavior is preserved. The canonical append sequence and
source-order merge remain authoritative after lane execution.

Unknown and recovered canonical inputs remain conservative. For complete
canonical procedure IR, an unresolved external call without local HTTP syntax
may skip the HTTP lane because the procedure-local HTTP solver has no local
state transition for that call; recovered documents, unknown CFG flow,
recovered or conditional statements/expressions, missing canonical inputs, and
explicit HTTP evidence, including sensitive logging or constants, keep the lane
planned.

Plans execute inside the existing bounded procedure-worker budget. The analyzer
does not create rule-level goroutines, per-diagnostic workers, or nested
semantic-kernel pools. Kernel and projection work accepts the analysis context
and checks cancellation at long-running work boundaries. Cancellation stops
queued or active work and publishes no partial final findings.

The performance recorder exposes additive, stderr-only counters:
`analysis_plans` (procedure plans executed), `planned_kernel_runs` (kernels
remaining after dependency closure), `skipped_kernel_runs` (enabled kernels
proven irrelevant), and `semantic_results_reused` (additional projections
reading an already materialized immutable result). These counters never appear
in normal analysis JSON or LSP diagnostics and do not change public
configuration or API contracts.

Array participant planning additionally reports
`array_participant_procedures`, `array_interprocedural_cfg_walks`, and
`array_worklist_revisits`. These counters describe one revision's semantic
closure and fixed-point work; they are not findings and are not a public IR,
CLI, JSON, or LSP field. The participant counter is a unique closure size,
the interprocedural CFG counter counts walks actually started by array
summaries/entry propagation, and the revisit counter excludes each
participant's initial evaluation.

Data-flow lane telemetry additionally exposes
`planned_generic_dataflow_runs`, `skipped_generic_dataflow_runs`,
`planned_http_dataflow_runs`, `skipped_http_dataflow_runs`,
`generic_dataflow_kernel_runs`, `http_dataflow_kernel_runs`,
`generic_dataflow_cfg_walks`, and `http_dataflow_cfg_walks`. Existing aggregate
`dataflow_cfg_walks` and `semantic_kernel_runs` counters remain compatibility
measurements for the procedure execution and are not lane sums.

### Indexed semantic-state solver boundary

The HTTP transport kernel and the ordinary Array block-level walk use the
internal `internal/analyze/semanticstate` solver. The ordinary Array adapter
uses compact flow slots; HTTP currently uses one compatibility slot containing
the legacy nested state while semantic-unit scalarization remains follow-up
work. A procedure revision builds
an immutable, case-insensitive environment with revision-local `SymbolID`
values and a deterministic CFG `BlockOrdinal` index. Mutable flow states,
scratch buffers, cancellation checks, and the `(BlockOrdinal, LaneOrdinal)`
worklist remain owned by one solver invocation; no state or index is shared
across revisions or nested worker pools.

The solver stores present slots in dense bitset-backed storage for small or
dense domains and sorted sparse storage for wide, sparse domains. Domain joins
update the destination in place and report changed slots, so unchanged blocks
are not requeued. Exceptional and uncertain edges retain predecessor input
state. HTTP and Array adapters keep their existing conservative lattice and
diagnostic projection contracts; source ranges, representative evidence, and
finding multiplicity are not part of semantic state.

This is an incremental migration boundary. Array source-line traversal,
edge-refinement paths, and interprocedural evidence walkers continue to use
their compatibility walker until their callback contracts can be represented
without changing VBA227/VBA208/VBA249 behavior. The shared solver API and
dense/sparse/hybrid benchmarks are internal implementation details and do not
add CLI, JSON, or LSP fields.

### Batch procedure-worker boundary

Batch `analyze` may evaluate procedures concurrently for a large file after
the project and file facts have been constructed. Procedure workers share the
same immutable `moduleAnalysisFacts`, `procedureAnalysisFacts`, resolution
snapshot, effect/object summaries, and analyzer indexes; flow state, rule
overlays, statistics deltas, and finding buffers remain procedure-local. The
file coordinator owns file-global one-time diagnostics and performs any
deduplication that depends on source order.

The batch scheduler uses one bounded execution budget for file jobs and
procedure batches. It does not create an independent full-size pool inside
each file worker, and ordinary small files use the direct loop path. Procedure
results are written to stable procedure-indexed slots, merged in source order,
and passed through the existing final sort and suppression stages. Cancellation
stops queued work, is checked by active context-aware scans, and does not
publish partial findings.

Workers consume owned IR, facts, and source values only. They must not retain
tree-sitter nodes, a `ParsedDocument`, or borrowed `ParsedView.Source` data;
all tree/source access remains within `ParsedDocument.Read`. This is an
execution boundary for batch analysis and does not add a public CLI, JSON, or
LSP contract.

## Compatibility and Layer Boundaries

Existing adapters may project IR facts into `calls.CallSite`, analyzer findings,
symbol results, or LSP-owned values. Those projections preserve:

- `inspect calls` JSON fields, caller/property metadata, resolution statuses,
  candidate meaning, and sort order;
- analyzer finding fields and ordering;
- `VBA204` cleanup-label handling, message, reason, suggestion, nearby code,
  severity, and inline suppression; and
- CLI byte-based and LSP UTF-16 diagnostic ranges.

Read-only resolution consumers should use `ResolveView` and its lookup helpers
so they do not depend on the storage location of project-dependent facts.
Consumers that retain or publish an independently owned resolved snapshot may
continue to use materialized `Resolve`. This overlay boundary is internal and
does not add a public CLI flag, configuration key, JSON field, diagnostic ID, or
LSP capability.

Issue #426 intentionally stops at procedure syntax and conservative name/call
resolution. The separate CFG layer defined by
`docs/specs/vba-control-flow-graph.md` consumes this IR to provide basic blocks,
edges, reachability, and path queries without changing the meaning of the IR's
syntactic `Blocks`. `docs/specs/vba-effect-analysis.md` consumes resolved IR and
CFG facts to add direct and transitive effect summaries plus fixed-point
propagation without adding effect claims to the IR itself.

The IR itself does not expose reachability or effect claims. No new public CLI
flag, configuration key, JSON field, diagnostic ID, or LSP capability is
introduced by this specification.
