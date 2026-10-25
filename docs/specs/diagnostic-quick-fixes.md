# Diagnostic Quick Fixes

## Edit ownership

The Go intelligence layer generates document-local actions. Each action has a
title, LSP-compatible kind, an ordered list of non-overlapping edits, and
optional diagnostic associations. Every range uses UTF-16 editor coordinates
and refers to the input document revision. Adapters must project every edit
without applying them sequentially against changing coordinates.

The LSP retains request revision/lifecycle validation and emits versioned
`documentChanges` for clients advertising that capability. Other clients use
the existing `changes` fallback. The adapter projects associated diagnostics
without running diagnostics again. Opening an unsaved buffer makes that
buffer authoritative. Source edits use the editor's normal Undo mechanism.

## Discovery

Filter `context.only` before computing action families. If all correction
providers are disabled, a quickfix-only request does not parse the document.
Enabled providers borrow the existing parsed snapshot and scan only their
rule. TypeDB-backed providers remain enabled when the DB is available and
complete, independently of configurable syntax rules. They must not call complete lint or construct ProcedureIR/CFG to find
corrections. Apply existing rule-specific inline suppressions and match the
requested selection against the diagnostic or edit ranges.

`fix_available` indicates that a rule has a correction provider. It does not
guarantee a correction for each occurrence, enable a diagnostic, grant Fix All
eligibility, or classify suppression comments as corrections.

## Initial corrections

| Rule  | Correction                                 | Preconditions                                                                                                                |
| ----- | ------------------------------------------ | ---------------------------------------------------------------------------------------------------------------------------- |
| VB044 | Replace the procedure-name string constant | Existing configured constant-name, procedure, literal, and suppression contract                                              |
| VB084 | Remove redundant `Step 1`                  | Enabled rule, unrecovered CST, literal step text `1`, verified Step token, horizontal whitespace throughout the removed span |
| VB080 | Remove optional `Let`                      | Enabled rule, unrecovered CST, verified keyword and horizontal trivia before the assignment target                           |
| VB086 | Remove redundant `Option Base 0`           | Enabled rule, unrecovered CST, existing diagnostic predicate, single physical line                                           |
| VB001 | Insert `Option Explicit`                   | Enabled rule, unrecovered CST, proven absence and a safe module declaration boundary                                         |

VB084 removes the clause and immediately preceding horizontal whitespace.
Keep the following comment, statement separator, body, and original EOL.
Continuations elsewhere in the loop header remain untouched. If a continuation
occurs immediately before `Step` or between `Step` and `1`, offer no correction.
Expressions such as `(1)`, `+1`, or `1 + 0` are outside the existing diagnostic
predicate and must not be generalized by the correction provider.

VB080 removes the keyword and horizontal whitespace before the assignment
target. A continuation after the assignment operator is preserved, but a
continuation immediately after `Let` withholds the correction. Property Let
declarations and ordinary assignments are outside this provider.

VB086 removes only the statement's CST span, retaining indentation, comments,
and the physical line ending. A now-blank line is intentional. It does not
remove `Option Base 1` or a statement containing physical continuations.

VB001 is an explicit policy correction: it may reveal undeclared-variable
errors. It must never be treated as a behavior-neutral transformation or as
eligible for initial Fix All support. Preserve export headers, module
attributes, source encoding, and existing line endings at insertion.
Insert immediately after the leading export header/module attributes, before
existing options or following comments, so documentation stays attached to
its declaration. With no export preamble, insert at byte zero. Infer the
inserted EOL from the first existing line ending, defaulting to LF for an empty
document. Refuse late/unsafe header placement and ambiguous conditional or
continued `Option Explicit` spellings. Existing BOM rejection remains unchanged.

VB001 retains its existing default-enabled warning policy. VB044 remains
opt-in; VB080, VB084, and VB086 remain opt-in at information severity.
Documentation generation retains its existing refactor kind and has no
associated diagnostic. CLI application is defined below. Save-time fixes and
cross-file edits remain outside this contract.

## Document Fix All

VB044, VB080, VB084, and VB086 have explicit document Fix All eligibility,
independent of `fix_available`. VB001 and VB036/VB037 remain individual-only.
The rule's existing configuration and conservative correction preconditions
still apply; occurrences without a safe individual correction are unchanged.

The kinds are `source.fixAll.xlflow.vb044`, `source.fixAll.xlflow.vb080`,
`source.fixAll.xlflow.vb084`, and `source.fixAll.xlflow.vb086`. An explicit
`context.only` request for a kind or its `source`, `source.fixAll`, or
`source.fixAll.xlflow` parent requests aggregation. Parent requests return
separate actions per rule, never one mixed-rule transaction. Empty `only`, an
empty kind, and ordinary `quickfix` requests do not trigger aggregation.
Automatic requests (`triggerKind: 2`) omit source actions even when a source
kind is supplied. Save-time fixes are not enabled automatically.

Reuse the same syntax correction providers, suppression rules and diagnostic
projection as individual actions, with the entire captured document as the
selection. Do not run whole-file lint, TypeDB queries, semantic indexing or
ProcedureIR/CFG construction. Disabled or unsupported kinds return before
parsing. Sort edits by UTF-16 start/end positions and validate every range
against that revision, including surrogate boundaries. Duplicate ranges,
overlaps, insertions sharing a start, reversed ranges and invalid coordinates
reject the entire affected rule action. Adjacent ranges are permitted.

Publish Fix All only for open documents and clients advertising
`workspace.workspaceEdit.documentChanges`. Advertise the four source kinds
only to those clients. One action contains one versioned `TextDocumentEdit`
with all rule edits and all associated diagnostics. Request cancellation,
document changes, close and reopen invalidate stale results through the
existing bounded request lifecycle. Legacy clients retain individual fixes
through `changes`; they do not receive unversioned Fix All. Editor application
uses the normal workspace-edit and Undo mechanism.

## CLI application

`xlflow fix <path> --rule VB044|VB080|VB084|VB086 [--apply] [--json]` runs
same-rule Fix All for one file. It requires `xlflow.toml` in the current
project. Eligible files are under the configured modules, classes, or
workbook source roots, or the project test root. The initial CLI surface
accepts only `.bas` and `.cls`. Under the configured forms root, the sole
exception is `forms/code/*.bas` when `[userform].code_source = "sidecar"`;
forms assets, specs, and inactive sidecar paths are rejected even when a file
has a supported extension. UserForm `.frm` files are excluded because their
`.frx` sidecars and source ownership require a separate transaction contract.

The command reuses the shared `intel` same-rule correction provider, including
its enabled-rule configuration, inline suppression behavior, diagnostic
associations, and UTF-16 edit application. It does not enable rules, broaden
diagnostic defaults, run whole-project lint, load TypeDB, or construct
ProcedureIR/CFG. It reads the saved file from disk; open editor buffers,
workbooks, Excel COM, and workbook sessions are not inputs.

Without `--apply`, the command previews the proposed result. With `--apply`, it
captures and analyzes the current file again; a prior preview is never stored
or reused as an edit plan. The input must be valid UTF-8 without a BOM.
`changed` reports whether applying the proposed edits changes the captured
bytes. `applied` is true only when changed bytes were successfully written.
Preview output always has `applied=false`; an apply with no byte changes also
has `applied=false`.

JSON success uses the normal envelope with `command="fix"` and an `output`
object containing:

- `mode`: `preview` or `apply`, according to the requested operation.
- `path`: project-relative path with `/` separators.
- `rule`: uppercase rule ID.
- `changed` and `applied`: booleans with the meanings above.
- `source_sha256` and `result_sha256`: SHA-256 of the captured source bytes and
  the proposed result bytes. In preview mode they bind the result to that
  invocation's current file capture; they do not create a reusable plan.
- `edits`: ordered edits with ranges in the captured document revision. Each
  range uses zero-based UTF-16 line and character positions and each edit has
  `new_text`.
- `diagnostics`: the exact diagnostic metadata associated with the shared
  corrections.
- `diff`: unified diff from captured source to proposed result; it is empty
  when there is no change.

Before publishing a changed file, the command stages a replacement in the same
directory and uses the platform's native atomic replace. It rejects symlink
targets and non-regular files. Apply mode rejects a read-only target only when
the proposed bytes differ. Before publication it rechecks the target identity,
mode, and contents against the captured file; a detected change fails with
`fix_source_changed`. The transaction removes its staging file on cancellation
or failure before publication. This is best-effort conflict detection: it does
not provide a race-free compare-and-swap against arbitrary external writers or
crash durability.

The command advertises a source-tree mutation capability to consumers through
policy `ResourceSourceTree`, operation `mutate`, and `ParallelSafe: false`.
This is coordination metadata only: no source-tree lock or workbook lease is
acquired, and there is no `--wait` behavior. File rechecking remains
best-effort against external writers.

| Error code                | Exit | Meaning                                                                                                  |
| ------------------------- | ---: | -------------------------------------------------------------------------------------------------------- |
| `config_error`            |    2 | Required project configuration is missing or invalid, using the standard configuration error projection. |
| `fix_args_invalid`        |    2 | Invalid command arguments, rule, path scope, or unsupported file extension.                              |
| `source_encoding_invalid` |    1 | The selected source is not UTF-8 without BOM, using the existing source-encoding projection.             |
| `fix_source_changed`      |    1 | Target identity, mode, or contents changed before publication.                                           |
| `fix_edits_invalid`       |    1 | The shared edit set failed range or conflict validation.                                                 |
| `fix_failed`              |    3 | File-system or other operational failure.                                                                |
| `fix_cancelled`           |    3 | The operation was canceled before publication.                                                           |

## Type-backed Set corrections

VB036 offers **Insert Set for object assignment**; VB037 offers **Remove Set
from value assignment**. Both require an unrecovered CST, a complete TypeDB,
and an unambiguous explicitly declared local, parameter or module variable in
the current source. Array, function-return, property/member and conditional
targets are excluded. Unknown, inferred-only, Object and Variant types cannot
justify an edit. An identifier RHS must also bind to one such declaration.
Local and parameter bindings must lie within the containing procedure's
declaration range. Property Get/Let/Set accessors sharing a name retain
separate scopes; duplicate declarations inside one accessor remain ambiguous.

VB036 accepts only a typed object identifier RHS whose assignment compatibility
is confirmed by TypeDB. It inserts `Set ` or replaces an explicit `Let` prefix.
Withhold it when either operand has named or flagged default-member metadata,
including ambiguous defaults. Known project types are also withheld when their
default-member metadata cannot prove reference intent. Project classes absent
from TypeDB cannot pass the required assignability proof; code actions must not
build ProcedureIR/CFG merely to populate project default-member maps.
VB037 requires a scalar RHS identifier or scalar literal and removes only
`Set` plus horizontal trivia. Calls, member expressions, New expressions,
default-property access and other value expressions are deferred. Existing
diagnostic predicates still decide whether an action can be published.

Only selected physical assignment lines receive type-backed diagnostic
evaluation. Reject labels, colon-separated statements, inline If, conditional
compilation and physical continuations. Comments, indentation and EOL remain
unchanged. VB036 follows normal suppression; VB037 remains non-suppressible.
Associated diagnostics retain the exact published metadata and range.

VB036 is an explicit object-reference intent correction, not a promise of
runtime equivalence: adding Set changes the assignment semantics. Removing
Set does not guarantee that unrelated conversions or runtime errors disappear.
Neither correction is eligible for Fix All in this stage.

## Verification

Focused tests cover exact output and unchanged non-target bytes, LF/CRLF,
UTF-16 coordinates, comments, colon-separated statements, safe and withheld
continuations, malformed syntax, suppression, disabled rules, idempotency,
multiple edits, and document ownership/version projection. Snapshot counters
protect discovery from accidental ProcedureIR/CFG construction.
Fix All tests also compare individual edits and diagnostic metadata, verify
idempotency, kind hierarchy and eligibility, reject conflicting or invalid
ranges, and exercise automatic-request filtering, client capability gating,
version ownership, cancellation and document lifecycle changes over JSON-RPC.
Aggregation introduces no new VBA transformation semantics and reuses the
existing local VBE evidence below.

The optional native editor check uses the actual extension LanguageClient and
CLI LSP to request VB080 Fix All, apply both edits, and restore the saved buffer
with one VS Code Undo. Build the CLI from the repository root with
`rtk powershell -NoProfile -ExecutionPolicy Bypass -File ./scripts/dev/go.ps1 build -o .tmp-quickfix/xlflow.exe ./cmd/xlflow`,
then run `rtk pnpm --dir editors/vscode test:fix-all`.
`XLFLOW_VSCODE_DEBUG_PATH` can select another built CLI;
`VSCODE_EXECUTABLE` can select an editor executable. This developer check does
not require Excel or replace the ordinary extension test suite.

TypeDB-dependent LSP tests use explicit temporary manifest-backed fixtures,
including both complete and incomplete states. Default-member withholding tests
include a reference-only positive control so missing metadata cannot make a
negative test pass without exercising the provider. Installed TypeLib caches
are not test evidence.

The local VBE probes `quickfix-step-one-before` and
`quickfix-step-one-after` observe the original and corrected syntax after the
known accepted/rejected controls. They are maintainability observations, not
compile-equivalent bindings for VB084. Compilation does not establish runtime
equivalence; the correction removes only VBA's literal default step.

The sequential `quickfix-local-before` / `quickfix-local-after` VBE pair covers
the combined VB001/VB080/VB086 edits on a module whose variables are declared.
These are policy observations, not compile-equivalent bindings or a promise
that adding Option Explicit preserves compilation for undeclared variables.

For Set corrections, `quickfix-set-object-before` / `quickfix-set-object-after`
observe VB036 syntax on explicit Worksheet references and remain policy
evidence. Existing `set-scalar-target`, `scalar-assignment` and
`set-object-target` controls protect VB037's compile-error boundary. Focused
tests compare correction output with these sources and published diagnostics.
