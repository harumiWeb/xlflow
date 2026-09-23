# VBA Call-Graph Reachability

## Scope

The opt-in `VB021` rule reports procedures that have no host-facing entry
surface and are definitely unreachable from the known project roots. The
reportable set covers `Private` and `Friend` procedures, plus `Public`
procedures inside a host-hidden module — a standard or class module that
declares `Option Private Module` and does not set `VB_Exposed = True`. It uses
the project-wide call graph assembled by `internal/vba/callgraph` and the root
classification in `internal/vba/reachability`.

## Root classification

The root set is built before internal procedures are classified. It includes:

- the configured `[project].entry` when it resolves uniquely;
- public or implicitly public, argument-free `Sub` procedures in standard
  modules as confirmed macro roots;
- other public or implicitly public `Sub`, `Function`, and `Property`
  procedures in standard modules as possible API roots, because callers from
  Excel, worksheet formulas, or external VBA are not present in the project
  call graph;

- test procedures in standard modules, only while the module remains
  host-visible — a `Test*` procedure inside a host-hidden module is reportable
  like any other host-hidden public procedure;
- public members of `VB_Exposed = True` class modules as possible roots,
  because external clients can invoke them without a project caller;
- `Auto_Open` and `Auto_Close`;
- recognized `Workbook_*` and `Worksheet_*` host-event procedures in document
  modules; the prefix alone does not make an arbitrary helper an event;
- UserForm event procedures and control event procedures when form metadata is
  available; and
- procedures whose names match a `WithEvents` field callback in the same class.

`Friend` members and `Public` procedures in a host-hidden module are not
roots: `Friend` is callable only inside the project, and `Option Private
Module` removes the host-visible surface unless the module carries
`VB_Exposed = True`. A module file that cannot be read from disk fails open —
its public procedures keep their external classification and remain roots.

An entry that cannot be resolved exactly is matched against procedure-name
candidates as a possible root. Ambiguous roots are possible roots, never
confirmed roots. Possible roots propagate confirmed project call edges as
possible reachability, so private helpers behind an externally callable API
are not reported as definitely unreachable.

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

`VB021` is emitted once at most for each unreachable reportable declaration. A
connected component of the unreachable confirmed-edge graph is
reported as one cluster context on its representative diagnostic; the
declarations in the component still retain individual locations so existing
inline suppression remains line-based. Dynamic references never become graph
edges and therefore cannot create confirmed clusters. The representative's
existing optional `context` field contains the cluster names; this is part of
the existing issue JSON shape and does not replace the per-declaration issue
locations.

The diagnostic ID, warning severity, configuration key,
`detect_unused_private_procedures` opt-in, inline suppression behavior, and
public CLI JSON shape remain unchanged. Dynamic reference facts are internal Go
data only.

## Verification contract

Tests cover isolated procedures, private chains, recursion and diamond-shaped
confirmed graphs, configured and ambiguous roots, host events, UserForms and
`WithEvents`, static and unknown dynamic callbacks, dynamic calls from
unreachable procedures, and one-diagnostic-per-declaration cluster reporting.
