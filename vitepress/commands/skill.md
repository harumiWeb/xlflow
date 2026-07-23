# xlflow skill

Install the bundled xlflow skill for AI agent tools.

## Usage

```bash
xlflow skill install [--agent <provider>|--target <path>] [--force]
```

## Options and Arguments

| Option / argument    | Description                                       | Default  |
| -------------------- | ------------------------------------------------- | -------- |
| `install`            | Install the skill files.                          | required |
| `--agent <provider>` | Install to a known agent profile such as `codex`. | -        |
| `--target <path>`    | Install to an explicit directory.                 | -        |
| `--force`            | Overwrite an existing skill installation.         | false    |
| `--json`             | Return installation paths.                        | false    |

## Examples

```bash
xlflow skill install --agent codex
xlflow skill install --target .agents/skills/xlflow --force --json
```

## Notes

::: warning
Use either `--agent` or `--target`, not both.
:::

::: tip
For repository-local agent setup, commit the generated skill only if your team intentionally shares that workflow.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "skill install",
  "target": ".agents/skills/xlflow",
  "written": ["SKILL.md"]
}
```

## Related

- [AI agents](../ai-agents/)
- [recommended prompts](../ai-agents/recommended-prompts)

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow skill` when the task matches the command description above. For a goal-oriented workflow, start with the [How-to guides](../guides/) and return here for exact options.

## Prerequisites

Check the project configuration and run `xlflow doctor --json` before workbook-backed operations. Source-only commands can run without Excel; commands that read or mutate a workbook require Windows Excel and VBIDE access.

## What this command reads and changes

The command reads the inputs and configuration described in its syntax and examples. Treat source files, the saved workbook, and a live session as separate states; add `--session` when the live workbook is authoritative. Any mutation is reversible only when a backup or explicit session save boundary exists.

## Effect on source-of-truth state

Use `xlflow status --json` before and after the command. A source edit normally requires `push`; a workbook edit normally requires `pull`; a dirty live session requires `save --session` or an intentional discard.

## Common workflows

Combine this command with the relevant [source/workbook/session workflow](../concepts/workbook-session-source), and use `--json` in scripts and agent loops.

## Common failures

Read the structured `error.code`, exit code, and recovery metadata instead of scraping terminal text. The [symptom-oriented troubleshooting guide](../help/troubleshooting) maps installation, execution, session, VS Code, and WSL failures to recovery steps.
