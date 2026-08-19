<!-- xlflow-rule-contract: {"id":"VB052","family":"lint","category":"correctness","default_severity":"error","scope":"interprocedural","realtime":true,"inline_suppressible":false,"preflight_blocking":true} -->
<!-- xlflow-rule-contract: {"id":"VB053","family":"lint","category":"correctness","default_severity":"error","scope":"interprocedural","realtime":true,"inline_suppressible":false,"preflight_blocking":true} -->
<!-- xlflow-rule-contract: {"id":"VB054","family":"lint","category":"correctness","default_severity":"error","scope":"interprocedural","realtime":true,"inline_suppressible":false,"preflight_blocking":true} -->

# VBA Resolution Diagnostics

`procedureir` is the canonical project resolver for calls, enum members, and
`RaiseEvent` references. It is shared by batch lint/analyze and the LSP
workspace snapshot; `calls.Resolver` only provides the legacy inspect-calls
projection.

The resolver distinguishes matched, non-callable, ambiguous, unresolved,
external, built-in, dynamic, and incomplete outcomes. `VB052` reports only a
provably project-local missing or non-callable call target. Bare unresolved
names remain quiet because they may bind to a referenced library or late-bound
value. `VB053` reports only an unqualified Enum member with multiple visible
project declarations and no same-module lexical winner. TypeLib enum metadata
is treated as an external fallback: duplicate records for one globally exposed
constant do not establish a source-level ambiguity. `VB054` reports an event
name that is not declared in the same object module as the `RaiseEvent`
statement.

All three rules are compile-equivalent errors, block source preflight, and are
not inline-suppressible. They fail open for parser recovery, unresolved
conditional branches, partial/path-filtered projects, pending or failed LSP
snapshots, unavailable TypeLib data, external members, late binding,
`Application.Run`, `CallByName`, and dynamically constructed names. LSP Fast
diagnostics never publish these project-negative outcomes; Full diagnostics
retry after a complete workspace snapshot is available.
