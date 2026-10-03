# ADR-0062: Unified Project Architecture Report

## Status

Accepted

## Context

Issue #463 asks for one deterministic snapshot to orient an agent or reviewer
in an unfamiliar VBA project. Existing symbols, calls, dependencies,
reachability, metrics, hotspots, module-state facts, and procedure effects
already answer individual questions. Running their CLI commands sequentially
would repeat discovery, parsing, and resolution and could combine different
source revisions.

Hotspot aggregation currently belongs to the metrics CLI. Module-state facts
are collected alongside optional VBA240 findings. Those ownership boundaries
prevent a reporting consumer from reusing facts without importing command or
diagnostic policy.

## Decision

Add a dedicated source-only `architecture` command over a protocol-neutral
architecture projection. Read each selected project source once, retain its
parsed document and IR/CFG, and use the canonical resolver, graph,
reachability, effects, metrics, and hotspot implementations. Keep fact
collectors separate from diagnostic selection and threshold enforcement.

Move hotspot aggregation into its owning package and expose the existing
module-state collector independently from VBA240 emission. Shared collectors
must preserve existing consumers' meanings and compatibility contracts.

Resolve the entire project before applying `--module` or `--path` display
filters. Maintain project-wide roots, reachability, fan-in/out, and hotspot
ranks. Retain incident dependency edges and complete related SCCs with
explicit boundary entities. A display filter does not establish absence of a
callee or turn a procedure into unreachable code.

The architecture view ignores metrics-only exclusions and threshold selectors.
Its versioned schema keeps canonical confirmed dependencies separate from
uncertain relationships and preserves nested schema/model versions. Human
output is bounded; JSON keeps the complete selected view and its evidence.

## Consequences

- Agents and Linux CI can obtain one architecture snapshot without Excel,
  COM, VBA execution, project-state writes, or enabling diagnostics.
- Filtered reports retain valid whole-project context but cost a whole-project
  parse; they are presentation filters rather than performance shortcuts.
- Architecture and metrics can rank different cohorts when metrics exclusions
  are configured. The score model is shared, while the cohort is explicit.
- Existing hotspot Excel-reference counts remain signals; detailed direct
  effects use the canonical effects model. Neither is proof of runtime
  execution or a new diagnostic gate.
- Shared extraction requires equivalence tests and analyzer performance checks
  so a new report does not broaden ordinary diagnostic work.

## Alternatives Considered

1. **Compose CLI outputs.** Rejected because it repeatedly scans the project
   and cannot guarantee one coherent revision.
2. **Add `analyze architecture`.** Rejected because runtime-risk diagnostic
   policy is a different responsibility from informational reporting.
3. **Resolve only filtered sources.** Rejected because it loses callers,
   roots, and dependency targets; it would require a weaker, incomplete model.
4. **Redefine hotspot scores or infer dynamic targets.** Rejected because
   consumers need reusable canonical facts with explicit uncertainty.

## Related

- Issue #463; related issues #458–#462
- ADR-0021, ADR-0025 (rooted callgraph), ADR-0036 (metrics), ADR-0039, ADR-0048
- `docs/specs/architecture-command.md`
- `internal/vba/architecture`, `internal/vba/hotspots`, `internal/vba/reachability`
- Architecture CLI and collector regression tests
