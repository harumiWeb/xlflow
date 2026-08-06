# VBA Call-Graph Reachability

## Scope

The opt-in `VB021` rule reports only private procedures that are definitely
unreachable from the known project roots. It uses the project-wide call graph
assembled by `internal/vba/callgraph` and the root classification in
`internal/vba/reachability`.

## Root classification

The root set is built before private procedures are classified. It includes:

- the configured `[project].entry` when it resolves uniquely;
- public or implicitly public, argument-free `Sub` procedures in standard
  modules;
- test procedures in standard modules;
- `Auto_Open` and `Auto_Close`;
- `Workbook_*` and `Worksheet_*` procedures in document modules;
- UserForm event procedures and control event procedures when form metadata is
  available; and
- procedures whose names match a `WithEvents` field callback in the same class.

An entry that cannot be resolved exactly is matched against procedure-name
candidates as a possible root. Ambiguous roots are possible roots, never
confirmed roots.

## Confirmed and possible reachability

Only a `matched` call with exactly one project-local candidate becomes a
confirmed graph edge. Confirmed roots and their confirmed edges are propagated
to a fixed point.

Known dynamic APIs are extracted separately from ordinary calls:

- `Application.OnTime` (`Procedure` or positional callback argument);
- `Application.OnKey` (positional callback argument);
- `Application.Run` (macro name argument); and
- `CallByName` (method-name argument).

The extractor preserves the source expression and argument metadata internally
without changing `inspect calls` JSON. Quoted strings and top-level string
concatenations are folded to static targets. Named arguments are preferred
when an API defines a named callback parameter; unresolved expressions retain
an unknown target.

A static dynamic target makes the target possibly reachable. An unknown dynamic
target from a confirmed or possibly reachable caller makes every project
procedure possibly reachable, because VBA can select any private procedure at
runtime. Dynamic references from an unreachable procedure do not suppress
`VB021`. Possible procedures remain distinct from confirmed procedures and are
never reported as definitely unreachable.

## Reporting

`VB021` is emitted once at most for each unreachable private declaration. A
private-only connected component of the unreachable confirmed-edge graph is
reported as one cluster context on its representative diagnostic; the
declarations in the component still retain individual locations so existing
inline suppression remains line-based. Dynamic references never become graph
edges and therefore cannot create confirmed clusters.

The diagnostic ID, warning severity, configuration key,
`detect_unused_private_procedures` opt-in, inline suppression behavior, and
public CLI JSON shape remain unchanged. Dynamic reference facts are internal Go
data only.

## Verification contract

Tests cover isolated procedures, private chains, recursion and diamond-shaped
confirmed graphs, configured and ambiguous roots, host events, UserForms and
`WithEvents`, static and unknown dynamic callbacks, dynamic calls from
unreachable procedures, and one-diagnostic-per-declaration cluster reporting.
