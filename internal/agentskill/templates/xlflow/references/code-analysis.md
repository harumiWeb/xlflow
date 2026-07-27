# Structural Code Analysis Reference

Load this reference before changing existing behavior when a procedure or
property's signature, visibility, name, or location may change; when deleting or
substantially refactoring code; when working in shared utilities; or when callers,
callees, dependencies, architecture, blast radius, or affected tests are unclear.

Structural analysis is source-only and read-only. Use it to narrow source review
and test selection, not to replace behavioral verification.

## Choose a Bounded Query

Start with the smallest question that can guide the next edit:

- Use `xlflow impact Module.Procedure --direction callers --json` to ask what
  project-local behavior may break if that procedure changes.
- Use `xlflow impact Module.Procedure --direction callees --json` to understand
  what the behavior depends on before changing it.
- Use `xlflow graph dependencies --module ModuleName --json` to understand local
  module, type, and implementation coupling around a shared area.
- Use `xlflow inspect calls --from Module.Procedure --json` when individual call
  syntax needs closer review.

Do not dump the full project graph into context by default. Widen depth or scope
only when the initial evidence shows that it is necessary.

## Interpret Results Conservatively

Use confirmed incoming dependencies to estimate blast radius and choose tests.
Use confirmed outgoing dependencies to understand prerequisites, collaborators,
and data flow.

Resolved nodes and edges are evidence, not a complete model of VBA runtime
behavior. Treat unresolved, ambiguous, external, built-in-like, member, and other
uncertain edges as gaps in static knowledge. An empty resolved caller list does
not prove that a signature change is isolated when uncertainty remains.

Follow source locations from the results to inspect only the relevant modules,
procedures, and tests. Resolve or account for uncertainty before making a
high-impact API change; do not silently assume it has no effect.

## Refactor Loop

1. Identify the changed symbol or module and run a bounded incoming or outgoing
   query before editing.
2. Use confirmed evidence and uncertainty to choose the source to inspect,
   focused tests, and broader verification scope.
3. Make the smallest refactor that preserves the intended contract.
4. Run lint, analysis, and focused behavior checks through the normal proof loop.
5. Re-run structural analysis when the refactor changed dependencies, names,
   visibility, locations, or module/type relationships; then run broader affected
   verification.

Skip graph analysis for obviously isolated changes such as comments, formatting,
or unrelated workbook styling.
