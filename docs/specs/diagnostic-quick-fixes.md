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
rule. They must not call complete lint or construct ProcedureIR/CFG to find
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

VB084 removes the clause and immediately preceding horizontal whitespace.
Keep the following comment, statement separator, body, and original EOL.
Continuations elsewhere in the loop header remain untouched. If a continuation
occurs immediately before `Step` or between `Step` and `1`, offer no correction.
Expressions such as `(1)`, `+1`, or `1 + 0` are outside the existing diagnostic
predicate and must not be generalized by the correction provider.

VB044 remains opt-in, and VB084 remains opt-in at information severity.
Documentation generation retains its existing refactor kind and has no
associated diagnostic. Fix All, CLI mutation, save-time fixes, and cross-file
edits are outside this initial contract.

## Verification

Focused tests cover exact output and unchanged non-target bytes, LF/CRLF,
UTF-16 coordinates, comments, colon-separated statements, safe and withheld
continuations, malformed syntax, suppression, disabled rules, idempotency,
multiple edits, and document ownership/version projection. Snapshot counters
protect discovery from accidental ProcedureIR/CFG construction.

The local VBE probes `quickfix-step-one-before` and
`quickfix-step-one-after` observe the original and corrected syntax after the
known accepted/rejected controls. They are maintainability observations, not
compile-equivalent bindings for VB084. Compilation does not establish runtime
equivalence; the correction removes only VBA's literal default step.
