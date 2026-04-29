---
name: adr-manager
description: Unified ADR workflow for any repository. Use when Codex needs to determine whether a change needs an Architecture Decision Record, draft a new ADR or propose an update, lint or review an ADR draft, audit ADRs against docs/specs/tests/src for drift, or refresh ADR index artifacts in repositories that already define ADR indexing conventions. Trigger on requests about ADRs, architecture decisions, design rationale, issue/PR/diff policy changes, ADR review, ADR audits, or ADR index maintenance.
---

# ADR Manager

## Overview

Handle the full ADR lifecycle through one entry point. Detect repository-specific ADR conventions first, then route to the right mode and fall back to bundled defaults only when the repository does not define its own template or workflow.

## Mode Routing

Choose one mode before producing output:

- `decide`
  - Use for issue, PR, diff, review-thread, or design-change triage.
  - Return whether an ADR is `required`, `recommended`, or `not-needed`.
- `draft`
  - Use for creating a new ADR draft or proposing an update to an existing ADR.
  - If the change is not ADR-worthy, return `ADR not needed` instead of forcing a draft.
- `lint`
  - Use for structural validation of an ADR draft.
  - Focus on status, required sections, evidence quality, and supersede linkage.
- `review`
  - Use for design review after `lint` has no unresolved `high` or `medium` findings.
  - Focus on decision quality, conflicts, rollout risk, and human-owned decisions.
- `audit`
  - Use for drift checks between ADRs and the current repository state.
  - Compare ADR claims with docs, specs, tests, and source paths.
- `index`
  - Use only when the repository already has ADR index artifacts or an explicit ADR index spec.
  - Do not invent derived index files for repositories that do not already use them.

If the user request is ambiguous, infer the mode from the artifact:

- existing ADR file -> `lint` by default
- explicit request to review an existing ADR -> run `lint` first when no prior clean `lint` result is available, then continue to `review` only if there are no unresolved `high` or `medium` findings
- issue / PR / diff asking "do we need an ADR?" -> `decide`
- request to "write", "draft", or "update" an ADR -> `draft`
- request to compare ADRs with implementation or find drift -> `audit`
- request to refresh ADR map, README, or index -> `index`

## Repository Discovery

Before using bundled defaults, detect the repository's ADR convention.

Read in this order:

1. ADR governance, criteria, workflow, or contributing docs that mention decision records
2. ADR directories and templates
3. Public docs and internal specs
4. Existing ADRs
5. Tests
6. Source code
7. Issue / PR / diff context

Search likely ADR roots first:

- `docs/adr`
- `docs/decisions`
- `adr`
- `architecture/adr`
- `dev-docs/adr`

Search likely repo policy docs next:

- `adr-governance.md`
- `adr-criteria.md`
- `adr-workflow.md`
- `contributing.md`
- decision-record guidance under `docs/`, `specs/`, or `architecture/`

Prefer repository-native conventions when they exist:

- path layout
- template headings and section names
- allowed status values
- naming and numbering scheme
- index artifact names and contract
- documentation language

If the repository does not define these, fall back to the bundled template and contracts in `references/default-adr-template.md` and `references/output-contracts.md`.

## Common Rules

- Treat ADRs as decision records. Keep `why` in the ADR and leave detailed `how` in specs or code.
- Build evidence from at least one concrete source. Prefer the full triad: docs/specs, tests, and src.
- Match the repository's language and formatting style. If there is no signal, write artifact bodies in English.
- When the change touches public API, CLI, MCP, schema, migration, fallback, or compatibility policy, include the corresponding public docs in scope.
- Do not silently rewrite accepted ADR text during `lint`, `review`, or `audit`. Return findings or an update proposal.
- If the change is not ADR-worthy, say so explicitly and include a short rationale instead of drafting noise.
- Escalate decisions that AI should not settle alone, including security, license, legal/compliance, unresolved product direction, public breaking-change judgment, or major organization-wide restructures.

## Mode Contracts

### `decide`

Return:

- `verdict`
- `rationale`
- `affected domains`
- `existing ADR candidates`
- `suggested next action`
- `evidence triad`

Use `required`, `recommended`, or `not-needed`.

### `draft`

Use the repository's existing ADR template when present. Otherwise use `references/default-adr-template.md`.

Return one of:

- new ADR draft
- update proposal for an existing ADR
- `ADR not needed`

Include concrete evidence paths whenever the repository provides them.

### `lint`

Return findings first, ordered by severity.

Check:

- valid status value
- required sections exist
- evidence is concrete rather than generic
- consequences include tradeoffs, not only benefits
- supersede references are present and internally consistent when claimed

Use severity `high`, `medium`, `low`.

### `review`

Only run after the current draft has no unresolved `lint` `high` or `medium` findings.
If the user asks for `review` directly and no clean `lint` result is in evidence, perform `lint` first and treat unresolved `high` or `medium` findings as a blocker to `review`.

Return:

- `verdict`
- `scope`
- `findings`
- `open questions`
- `residual risks`

Use verdicts `ready`, `revise`, `escalate`.

### `audit`

Audit claim-level alignment between ADRs and the current repository.

Return:

- `scope`
- `findings`

Use finding types:

- `policy-drift`
- `missing-adr-update`
- `missing-evidence`
- `stale-reference`

Each finding should include a recommended action:

- `update-adr`
- `new-adr`
- `update-specs`
- `add-tests`
- `no-action`

### `index`

Only run if the repository already has ADR index artifacts or an explicit index spec.

Return:

- `updated artifacts`
- `added or changed ADR entries`
- `consistency findings`

If no repository convention exists, return a concise no-op result that states the index convention was not found.

## Review and Audit Heuristics

- If a mode or backend policy changes across multiple entry points, it leans ADR-worthy.
- If a default changes consumer-visible meaning, treat it as a likely ADR candidate.
- If the repository already has a relevant ADR, prefer updating or superseding it over creating overlapping policy.
- When compatibility, migration, fallback, or safety impact matters, require those consequences to be covered explicitly.
- For audits, compare claims against the current docs/specs/tests/src rather than relying on ADR text alone.

## References

- Read `references/default-adr-template.md` when no repository template or section contract is present.
- Read `references/output-contracts.md` when you need exact output envelopes for `decide`, `review`, `audit`, or `index`.
