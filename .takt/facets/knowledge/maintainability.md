# Maintainability

## Change Shape

- Prefer the smallest design that satisfies the current behavior and tests.
- Keep changes logically cohesive. Do not mix unrelated cleanup with feature or bug work.
- Use purpose-named helpers when the same operation would otherwise be scattered across call sites.
- Avoid flag arguments and broad conditionals that hide separate responsibilities.
- Do not add compatibility modes or legacy support unless the user request or existing contract requires it.

## Boundary Design

- Resolve config, paths, flags, provider choices, and environment differences at the boundary.
- Pass normalized internal values into core execution.
- Do not make lower layers reload global config or reinterpret raw CLI/env values.
- Keep input interpretation, execution, and output rendering as separate phases.

## Review and Documentation

- Durable decisions belong in ADRs when tradeoffs matter to future maintainers.
- Durable rules and CLI contracts belong in `docs/specs/`.
- Session progress, temporary verification notes, and incomplete hypotheses belong in `tasks/todo.md`.
- Lessons belong in `tasks/lessons.md` only when they prevent repeated mistakes.

## Testability

- Prefer behavior tests over string-presence tests when the logic can be executed cheaply.
- Inject clocks/tickers or poll with generous deadlines instead of relying on tight sleeps.
- Keep shell and Excel-dependent tests explicit about platform/tool requirements.
- Treat unit tests as insufficient for UserForm Designer or VBIDE overwrite contracts unless real Excel COM E2E covers the behavior.

## Red Flags

- A change requires grepping many unrelated call sites to understand one operation.
- New optional parameters exist but are never wired from callers.
- Tests pass while lint fails due to unused wrappers.
- Defensive branches protect conditions all callers already guarantee.
- Documentation removes README detail without making replacement pages self-contained.
