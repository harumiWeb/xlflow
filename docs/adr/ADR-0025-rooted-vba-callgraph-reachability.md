# ADR-0025: Rooted VBA Call-Graph Reachability for VB021

## Status

Accepted

## Context

`VB021` previously treated a private procedure as unused when no parsed call
site named it. That model confused caller absence with runtime reachability:
Excel can enter VBA through configured macros, document and UserForm events,
tests, `WithEvents` handlers, and callback APIs whose procedure name is carried
as a string. It also allowed an arbitrary underscore in a standard-module name
to suppress a finding, even though the name was not an event.

The project already has a shared procedure IR and conservative call resolver
(ADR-0021), plus a call-graph layer that exposes uniquely matched project-local
edges. The new analysis must preserve those contracts and must not turn
uncertain or string-based dispatch into false confirmed dependencies.

## Decision

Add a rooted reachability layer above the existing call graph. Build confirmed
roots from the configured entry point, argument-free standard-module macros
and tests, host events, UserForm/control metadata, and `WithEvents` callback
handlers. Treat other public or implicitly public standard-module procedures
(`Sub`, `Function`, and `Property`) as possible API roots because Excel,
worksheet formulas, and external VBA callers are outside the project graph.
Propagate only uniquely matched project-local calls as confirmed reachability
to a fixed point, and propagate the same edges from possible roots as possible
reachability.

Represent `Application.OnTime`, `Application.OnKey`, `Application.Run`, and
`CallByName` callback names as `DynamicReference` values separate from normal
call edges. Fold string literals and top-level literal concatenation when
possible. A static dynamic target is possible reachability, not confirmed
reachability. An unresolved dynamic expression from a reachable caller makes
all project procedures possible, since there is no sound candidate set for
late-bound VBA dispatch. An unresolved dynamic expression in an unreachable
caller has no effect on `VB021`.

If a configured entry is unresolved or ambiguous, matching procedure-name
candidates are possible roots rather than confirmed roots. A private procedure
is reported only when it is neither confirmed nor possible. This keeps the
rule conservative while still distinguishing proven graph reachability from
runtime possibility for future consumers.

Report every definitely unreachable private declaration at most once. Compute
private-only connected components over confirmed unreachable edges and attach
the component's names to one representative diagnostic through the existing
`Context` field. Individual declarations remain separate diagnostics so the
existing line-based inline suppression contract is preserved, while duplicate
path noise is avoided.

The dynamic-reference collection is internal and excluded from the existing
`inspect calls` JSON projection. The diagnostic ID, severity, configuration
key, opt-in behavior, and JSON issue shape remain unchanged.

## Consequences

- Host entry points and test procedures no longer depend on a textual caller
  being present in source.
- Public helper libraries, including scaffolded `Xlflow` modules, no longer
  produce false `VB021` findings for private helpers that are reachable from
  externally callable APIs.
- Dynamic callback targets avoid false positives without pretending to be
  confirmed graph edges.
- Unresolved dynamic calls can reduce `VB021` precision substantially, but only
  when their caller is itself reachable.
- Cluster context makes private call chains actionable without changing the
  number or location semantics of diagnostics.
- The root classifier must evolve when another host callback convention is
  supported; that policy remains centralized rather than embedded in linter
  naming heuristics.

## Alternatives considered

1. **Keep the caller-name check and add more underscore exceptions** -
   Rejected because naming conventions cannot identify configured entries,
   tests, host metadata, or string callbacks reliably.
2. **Convert dynamic references into confirmed call edges** - Rejected because
   a string expression and late-bound VBA call do not prove one target.
3. **Suppress every private procedure whenever any dynamic call exists** -
   Rejected because an unreachable procedure's dynamic call must not hide
   unrelated findings; suppression is scoped to reachable callers.
4. **Report only one issue for an entire cluster** - Rejected because it would
   break declaration-line inline suppression and make fixes less local.
5. **Exclude bundled `Xlflow` modules by path or module name** - Rejected
   because it would hide genuinely dead private code in helper modules and
   would not generalize to user-authored public libraries. Public API roots
   express the reachability reason without a framework-specific exemption.

## Evidence

- Shared IR and event classification: `internal/vba/procedureir`, ADR-0021.
- Conservative project edge construction: `internal/vba/callgraph` and ADR-0023.
- Dynamic extraction: `internal/vba/calls/calls.go`.
- Root construction: `internal/vba/reachability/reachability.go`.
- Consumer and compatibility behavior: `internal/lint/linter.go` and
  `internal/lint/linter_test.go`.
- Contract details: `docs/specs/vba-call-graph-reachability.md` and
  `docs/specs/cli-contract.md`.

## Related

- `docs/adr/ADR-0021-procedure-analysis-ir.md`
- `docs/adr/ADR-0023-procedure-effect-summaries.md`
- `docs/specs/vba-analysis-ir.md`
- Issue #458
