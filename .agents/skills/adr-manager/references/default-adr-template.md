# Default ADR Template

Use this template only when the repository does not already define its own ADR template.

```md
# ADR-0000: <Title>

## Status

`proposed` | `accepted` | `superseded` | `deprecated`

## Background

Why this decision is needed. Include constraints, context, and non-goals.

## Decision

What is being adopted. Include boundaries with public contracts if relevant.

## Consequences

Describe the positive and negative consequences of this decision.

## Rationale

- Tests:
- Code:
- Related specs:

## Supersedes

- None

## Superseded by

- None
```

Fallback defaults:

- Keep implementation details out of the ADR body unless they are needed to explain the policy boundary.
- Prefer concrete evidence paths over narrative claims.
- If the repository uses a different numbering or title convention, follow the repository instead of this example.
