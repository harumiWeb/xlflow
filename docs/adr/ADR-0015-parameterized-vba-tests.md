# ADR-0015: Parameterized VBA Tests

## Status

Accepted

## Context

xlflow's VBA test runner originally treated one public parameterless `Sub` as one test. That kept discovery and execution simple, but data-driven checks required duplicate procedures or manual loops inside one procedure. Manual loops hide the failing input combination behind one result and do not work well with exact filtering, fail-fast behavior, or workbook isolation.

Parameterized tests need to work in both source-only discovery (`test list`) and workbook-backed execution (`test`) without evaluating arbitrary VBA expressions during discovery.

## Decision

Introduce `'@TestCase(...)` metadata directly above public test procedures named `Test*` or `*_Test`.

Rules:

1. A parameterized test procedure must declare at least one `@TestCase`.
2. Each `@TestCase` expands to an independent executable test case with ID `<Module>.<Procedure>[<case>]`.
3. Named cases use `@TestCase("name"; args...)` and the name becomes the bracketed case ID.
4. Unnamed cases use a lexical canonical form of the literal arguments. Numeric semantic equivalence, such as `1.0` versus `1#`, is intentionally not normalized in the initial implementation.
5. Discovery accepts only safe scalar literals: integers, floating-point/scientific notation, strings, booleans, `Empty`, `Null`, and VBA date literals.
6. Initial parameter support is limited to `ByVal` scalar parameters with common VBA scalar types or `Variant`. `ByRef`, `Optional`, `ParamArray`, object parameters, arrays, and expression evaluation are rejected.
7. `@TestCase` validation failures use `error.code = "invalid_test_case"`.

The Go source discovery path and the C# .NET bridge do not share code, but they share the same model, validation rules, JSON shape, and regression fixtures so `test list --json` and `test --json` stay aligned.

## Consequences

- Positive: Repeated input/output checks get granular result identity and exact filtering.
- Positive: `--isolation test` naturally applies per case because the bridge already treats selected tests as the isolation unit.
- Positive: Discovery remains deterministic and does not execute user VBA or resolve workbook state.
- Negative: The annotation parser exists in both Go and C# and must stay in sync through tests and docs.
- Negative: Initial ID canonicalization is lexical, so some source formatting choices can still change unnamed case IDs.
- Negative: Advanced VBA values such as arrays, objects, constants, enum members, and workbook references require future explicit design.

## Related

- `docs/specs/cli-contract.md`
- `vitepress/commands/test.md`
- `docs/adr/ADR-0007-lifecycle-hooks.md`
- `docs/adr/ADR-0008-dotnet-excel-bridge.md`
