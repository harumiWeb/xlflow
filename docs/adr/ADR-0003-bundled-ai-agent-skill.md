# ADR-0003: Bundled AI Agent Skill

## Status

`accepted`

## Background

xlflow is intended to be used by AI agents that edit, run, test, and repair Excel VBA projects. The previous project scaffold created `prompts/agent.md`, but that file was too small to carry the full workflow contract and was tied only to newly initialized projects.

Alternative approaches considered:

- Keep generating `prompts/agent.md` during `new` and `init`.
- Publish the skill only through an external registry.
- Bundle an official `xlflow` skill in the CLI and let users install it into provider-specific project directories.

The skill needs to stay version-aligned with the CLI and work in offline or controlled Windows environments.

## Decision

Bundle the official `xlflow` AI agent skill in the CLI binary and install it through `xlflow skill install` or `xlflow new/init --with-skill`.

The supported provider targets are `.agents/skills`, `.codex/skills`, `.claude/skills`, `.cursor/skills`, `.gemini/skills`, and `.copilot/skills`. Non-interactive and JSON runs must provide `--agent` or `--target`; interactive runs may choose a provider through a Bubble Tea selector.

Stop generating `prompts/agent.md` from project scaffolding.

## Consequences

Positive consequences:

- The official agent workflow remains aligned with the installed xlflow version.
- Users can install the same skill during setup or later without relying on an external registry.
- Agent-specific targets are explicit, while CI and JSON callers stay deterministic.

Negative consequences:

- The CLI now depends on Bubble Tea for the interactive provider selector.
- Existing users who expected `prompts/agent.md` from scaffolding must install the skill instead.
- Updating customized project-local skills still requires a future update or diff workflow.

## Rationale

- Tests: `internal/agentskill`, `internal/cli`, and `internal/project` unit tests.
- Code: `internal/agentskill`, `internal/cli`, embedded `internal/agentskill/templates/xlflow`.
- Related specs: `docs/specs/cli-contract.md`, `docs/design.md`.

## Supersedes

- None

## Superseded by

- None
