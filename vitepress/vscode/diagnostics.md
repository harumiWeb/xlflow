# Diagnostics and navigation

Diagnostics run against the current in-memory editor buffer and publish stable `VB...`, `VBA...`, and UserForm codes with source `xlflow`. This means they can warn you before you save a file or open Excel. Use the Problems panel to jump to a line, hover for the explanation, and apply a Quick Fix where one is offered.

Registered `VB...` and `VBA...` diagnostics link from the Problems panel to the
matching entry in the generated [diagnostic catalog](../reference/diagnostics).
The extension reads rule metadata from `xlflow rules --json` and offers an
inline-suppression Quick Fix only when that installed xlflow version marks the
rule as suppressible. If metadata cannot be loaded, its schema is unsupported,
or the ID is unknown, the extension does not guess and omits the action.

Project-wide checks still run from `xlflow lint` and `xlflow analyze`; they include filesystem and preflight rules that cannot be inferred safely from one unsaved document. Treat a clean editor as a good first signal, then run the project checks before `push`. Use `Go to Definition`, `Find References`, symbols, and rename for project navigation.

Use **Source Action...** for **Fix all VB044/VB080/VB084/VB086 occurrences in this
document** when the corresponding lint rule is enabled. Each action corrects
one rule in the open buffer and can be undone normally. Unsafe occurrences
remain unchanged; conflicting edits withhold that rule's action. VB001 and
VB020 and VB036/VB037 remain individual Quick Fixes. VB020 offers removal only
for a standalone, single-variable `Dim` with an explicit primitive scalar type
and no reads or writes elsewhere in the procedure. Comments and line endings
are preserved. Fix All uses versioned document edits and is not enabled
automatically on save.

When `analyze.detect_byref_parameters_can_be_byval` is enabled, VBA275 can offer
**Pass parameter ByVal** for an explicit `ByRef` parameter in a private standard-
module Function. The first version requires one primitive scalar parameter and
scalar return-slot assignments using only that parameter and literals. Calls,
other bindings, public APIs, Optional, Variant, arrays and objects have no fix.
The action appears after current project-aware Full diagnostics complete; an
edit can temporarily hide it while new diagnostics are pending. Only `ByRef`
changes, preserving comments and line endings. VBA275 remains individual-only.
