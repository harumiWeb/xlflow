# ADR-0006: Tests Located Under `src`

## Status

Accepted

## Context

The current scaffold creates a top-level `tests/` directory, but `push` never imports it into the workbook. This makes `tests/` a dead folder that confuses users and breaks the "source of truth is always under `src/`" principle.

Rubberduck-style separate test directories also conflict with xlflow's headless/session-centered design because:

- They require separate import rules.
- They create a drift risk between `push` targets and `test` targets.

## Decision

Remove the `tests/` scaffold directory and recommend placing tests under `src/modules/Tests/` (or any `src/` subdirectory). Because `push` already enumerates source recursively, test modules under `src/modules/Tests/` are imported automatically with no extra configuration.

## Consequences

- Positive: `push` and `test` targets stay in sync naturally.
- Positive: Folder annotations (`@Folder`) can organize tests the same way as production modules.
- Negative: Users with existing `tests/` folders must manually move `.bas` files into `src/`.

## Related

- `docs/design.md`
- `docs/specs/cli-contract.md`
