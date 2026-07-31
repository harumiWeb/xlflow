# Diagnostics and navigation

Diagnostics run against the current in-memory editor buffer and publish stable `VB...`, `VBA...`, and UserForm codes with source `xlflow`. This means they can warn you before you save a file or open Excel. Use the Problems panel to jump to a line, hover for the explanation, and apply a Quick Fix where one is offered.

Registered `VB...` and `VBA...` diagnostics link from the Problems panel to the
matching entry in the generated [diagnostic catalog](../reference/diagnostics).
The extension reads rule metadata from `xlflow rules --json` and offers an
inline-suppression Quick Fix only when that installed xlflow version marks the
rule as suppressible. If metadata cannot be loaded, its schema is unsupported,
or the ID is unknown, the extension does not guess and omits the action.

Project-wide checks still run from `xlflow lint` and `xlflow analyze`; they include filesystem and preflight rules that cannot be inferred safely from one unsaved document. Treat a clean editor as a good first signal, then run the project checks before `push`. Use `Go to Definition`, `Find References`, symbols, and rename for project navigation.
