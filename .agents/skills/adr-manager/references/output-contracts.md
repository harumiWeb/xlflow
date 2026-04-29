# Output Contracts

Use these contracts only when the repository does not already define stricter ADR result shapes.

## `decide`

Return at least:

- `verdict`: `required` | `recommended` | `not-needed`
- `rationale`: 1-3 short lines
- `affected domains`
- `existing ADR candidates`
- `suggested next action`: `new-adr` | `update-existing-adr` | `no-adr`
- `evidence triad`
  - docs/specs
  - src
  - tests

## `review`

Return at least:

- `verdict`: `ready` | `revise` | `escalate`
- `scope`
- `findings`
- `open questions`
- `residual risks`

Each finding should include:

- `type`
- `severity`
- `summary`
- `why it matters`
- `suggested revision`
- `evidence`

Use finding types:

- `decision-gap`
- `scope-conflict`
- `evidence-risk`
- `rollout-gap`
- `ownership-escalation`

## `audit`

Return at least:

- `scope`
- `findings`

Each finding should include:

- `type`
- `severity`
- `claim`
- `affected ADRs`
- `evidence matrix`
- `recommended action`

Use finding types:

- `policy-drift`
- `missing-adr-update`
- `missing-evidence`
- `stale-reference`

Use recommended actions:

- `update-adr`
- `new-adr`
- `update-specs`
- `add-tests`
- `no-action`

## `index`

Return at least:

- `updated artifacts`
- `added or changed ADR entries`
- `consistency findings`

If the repository has no ADR index convention, say so explicitly and stop instead of inventing new artifacts.
