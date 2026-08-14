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
func Resolve(DocumentIR, Resolver) DocumentIR
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

`Resolve` returns an independently owned IR with a resolution overlay derived
from the supplied project snapshot. It does not mutate the input IR or rescan
source. Replacing the resolver can therefore re-resolve the same syntax after a
workspace symbol change.

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

- builds lazily from the snapshot-owned `ParsedDocument`;
- is safe for concurrent readers and performs at most one IR build per
  snapshot;
- shares the same parsed document already used by snapshot symbols,
  diagnostics, and calls;
- returns defensive copies so callers cannot mutate cached slices or nested
  candidate data; and
- is retired with the snapshot and never closes the parsed document itself.

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

## Compatibility and Layer Boundaries

Existing adapters may project IR facts into `calls.CallSite`, analyzer findings,
symbol results, or LSP-owned values. Those projections preserve:

- `inspect calls` JSON fields, caller/property metadata, resolution statuses,
  candidate meaning, and sort order;
- analyzer finding fields and ordering;
- `VBA204` cleanup-label handling, message, reason, suggestion, nearby code,
  severity, and inline suppression; and
- CLI byte-based and LSP UTF-16 diagnostic ranges.

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
