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
associated diagnostic. Fix All, CLI mutation, save-time fixes, and cross-file
edits are outside this initial contract.

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
