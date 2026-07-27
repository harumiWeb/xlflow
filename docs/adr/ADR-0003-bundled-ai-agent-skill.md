# ADR-0003: Bundled AI Agent Skill

## Status

`accepted`

## Background

xlflow is intended to be used by AI agents that edit, run, test, and repair Excel VBA projects. The previous project scaffold created `prompts/agent.md`, but that file was too small to carry the full workflow contract and was tied only to newly initialized projects.

Alternative approaches considered:

- Keep generating `prompts/agent.md` during `new` and `init`.
- Publish the skill only through an external registry.
- Bundle an official `xlflow` skill in the CLI and let users install it into provider-specific project directories.

The skill needs to stay version-aligned with the CLI and work in offline or controlled Windows environments. As xlflow's CLI, diagnostics, structured JSON, recovery, and analysis capabilities matured, the bundled guidance accumulated command reference material, feature workflows, and mechanically enforceable VBA rules alongside its agent decision-making rules. Keeping those concerns together increases context cost and risks stale or contradictory CLI guidance.

## Decision

Bundle the official `xlflow` AI agent skill in the CLI binary and install it through `xlflow skill install` or `xlflow new/init --with-skill`.

The supported provider targets are `.agents/skills`, `.codex/skills`, `.claude/skills`, `.cursor/skills`, and `.gemini/skills`. GitHub Copilot uses the shared `.agents/skills` target rather than a separate `.copilot/skills` target. Non-interactive and JSON runs must provide `--agent` or `--target`; interactive runs may choose a provider through a Bubble Tea selector.

Stop generating `prompts/agent.md` from project scaffolding.

Keep the installed `SKILL.md` as a thin orchestration contract: broadly applicable decision rules, safety invariants, one normal proof loop, a compact executable command sequence for normal session-first development, a purpose-to-command capability map, and dispatch rules for specialized work. The capability map directs unfamiliar agents to current CLI help and representative commands without becoming a command manual. Keep detailed workflows and domain knowledge in bundled `references/` documents. The CLI, its structured diagnostics, and official command documentation own complete syntax, flag inventories, error details, and mechanically enforceable checks.

The orchestration contract must direct agents to use structural analysis when an existing behavior may have callers, dependencies, or an uncertain blast radius. It must treat unresolved or ambiguous analysis as uncertainty rather than proof that no impact exists.

## Consequences

Positive consequences:

- The official agent workflow remains aligned with the installed xlflow version.
- Users can install the same skill during setup or later without relying on an external registry.
- Agent-specific targets are explicit, while CI and JSON callers stay deterministic.
- The main Skill has a smaller, stable context footprint while recovery, testing, forms, UI, debugging, formulas, and code-analysis workflows remain discoverable on demand.
- CLI behavior can evolve without duplicating flags, rule identifiers, or diagnostic schemas throughout agent guidance.

Negative consequences:

- The CLI now depends on Bubble Tea for the interactive provider selector.
- Existing users who expected `prompts/agent.md` from scaffolding must install the skill instead.
- Updating customized project-local skills still requires a future update or diff workflow.
- Agents must follow the Skill's reference dispatch rules when a task needs specialized guidance; the main Skill intentionally is not a complete CLI manual.

## Rationale

- Tests: `internal/agentskill`, `internal/cli`, and `internal/project` unit tests.
- Code: `internal/agentskill`, `internal/cli`, embedded `internal/agentskill/templates/xlflow`.
- Related specs: `docs/specs/cli-contract.md`, `docs/design.md`; public installation guidance in `vitepress/ai-agents/skills.md`.

## Supersedes

- None

## Superseded by

- None
