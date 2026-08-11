# VBA Procedure Call-Cycle Analysis

<!-- xlflow-rule-contract: {"id":"VBA244","family":"analyze","category":"reliability","default_severity":"information","scope":"project-wide","realtime":false,"configuration_key":"detect_procedure_call_cycles","inline_suppressible":true,"preflight_blocking":false} -->

`VBA244` reports confirmed recursive and cyclic dependencies between project
procedures. It runs in batch `analyze` and `check` and is enabled by default.
Disable it with `[analyze].disabled_rules = ["VBA244"]` or suppress an anchor
call with `xlflow:disable-line VBA244` / `xlflow:disable-next-line VBA244`.

## Graph contract

The detector reuses the project call graph. Only a uniquely resolved,
project-local call is a confirmed edge. Ambiguous, unresolved, external, and
dynamically bound calls never create or close a confirmed cycle. The detector
enumerates every elementary directed cycle, including self-recursion, mutual
recursion, cycles crossing modules, nested/chorded cycles, and independent
cycles.

Each cycle is canonicalized by rotating its ordered procedure path so the
lowest stable procedure identity is first. A reverse path is distinct when its
directed edges differ. Parallel call sites between the same procedures use the
earliest stable source location. Results and finding order are deterministic.

## Severity and context

An ordinary cycle is reported at `information`. It becomes a non-blocking
`warning` when a cycle procedure has an event-handler identity or a direct or
propagated reachable effect of Application-state mutation, error suppression,
workbook acquisition, or VBA file acquisition. Uncertainty is retained as
context but never by itself escalates severity.

The JSON finding has a `call_cycle` object containing the closed procedure
`path`, ordered `edges`, `cross_module`, event-handler identities, dangerous
effect evidence, and unresolved-call uncertainty. The finding location is the
first outgoing call in the canonical cycle rotation; one finding represents
one cycle.

`path` nodes contain `qualified_name`, `module`, `kind`, `module_kind`, `file`,
and declaration `line`. The final node repeats the first node. Each `edges`
entry contains `caller`, `callee`, and the source `file`, `line`, `column`,
`end_line`, and `end_column`; edge `i` leaves path node `i`. `event_handlers`
contains stable participating identities. `dangerous_effects` contains the
effect `kind`, originating procedure, source file and line, and an optional
normalized target. `uncertainty` contains the resolution `kind`, origin,
unresolved callee text when available, and call location. Empty optional
collections are omitted, so adding this object remains backward compatible for
consumers that do not understand `VBA244`.

The impact command's existing `cycles[].nodes` remains an open path for
backward compatibility; `cycles[].edges` is additive and aligns with the
ordered nodes, while `VBA244` uses the closed `call_cycle.path` form.

The cycle graph is structural and is not an all-path execution proof. Effect
evidence remains conservative and follows the existing procedure-effect
summary contract. `opens_file` records reachable, non-recovered VBA
`Open ... For ... As #...` acquisition statements; `VBA219` continues to own
resource ownership and leak diagnostics.
