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

::: important
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
