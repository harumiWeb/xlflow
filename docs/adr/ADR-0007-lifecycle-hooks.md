# ADR-0007: Lifecycle Hooks for VBA Tests

## Status

Accepted

## Context

Users need setup and cleanup around tests, but Rubberduck's annotation-driven `ModuleInitialize` / `TestInitialize` model is IDE-centric and does not map well to xlflow's headless CLI runner. We need a hook model that:

- Works without IDE annotations.
- Respects xlflow's "no hidden rules" principle.
- Keeps the runner simple and avoid private-method-invocation complexity.

## Decision

Introduce reserved public parameterless `Sub` names recognized per module:

- `BeforeAll`
- `AfterAll`
- `BeforeEach`
- `AfterEach`

Rules:

1. Only public (or implicit-public) parameterless `Sub` procedures are recognized.
2. Hooks are scoped to the module that declares them.
3. The runner executes them in this order:
   - `BeforeAll` once per module
   - For each test: `BeforeEach` → test body → `AfterEach`
   - `AfterAll` once per module
4. `BeforeEach` failure skips the test body but still runs `AfterEach` for cleanup.
5. `AfterEach` failure overrides the test result to `failed`.
6. `BeforeAll` or `AfterAll` failure marks **all** tests in the module as `failed`.
7. Hook failures use dedicated `error.code` values (`before_all_failed`, `after_all_failed`, `before_each_failed`, `after_each_failed`).
8. The generated runner keeps each test case in its own `RunTest_<index>` function so a growing suite does not exceed VBA's per-procedure size limit. The bridge invokes the selected function by name.

The runner generates a temporary VBA module with module-specific `RunBeforeAll_<ModuleName>`, `RunAfterAll_<ModuleName>`, and one `RunTest_<index>` function per executable test case. The bridge invokes the selected wrapper through Excel's `Run` entry point; each wrapper directly invokes module-qualified VBA procedures. This avoids private-dispatch complexity entirely while keeping each generated test procedure small.

## Consequences

- Positive: Simple, explicit, and CLI-friendly.
- Positive: No annotation parser or IDE dependency required.
- Negative: Hooks must be public, so they are visible to workbook VBA.
- Negative: The temporary module contains one wrapper per executable test, increasing the number of generated procedures, while avoiding the single-dispatcher size ceiling.

## Alternatives Considered

1. **Rubberduck annotation model** — Rejected because annotations are IDE metadata, not CLI metadata.
2. **Private hook invocation via VBIDE** — Rejected because it requires unsafe code injection and breaks the "no hidden rules" principle.
3. **Global single `BeforeAll`/`AfterAll`** — Rejected because module-scoped isolation is safer and easier to reason about.

## Related

- `docs/design.md`
- `docs/specs/cli-contract.md`
