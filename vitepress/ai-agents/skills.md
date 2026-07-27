# Skills

`xlflow skill install` installs the bundled `xlflow` Skill for agent tools. The Skill is a thin orchestration layer: it directs an agent through the normal proof loop and the safety decisions that apply to most Excel/VBA work, while xlflow itself remains the source of truth for command syntax, diagnostics, and enforcement. It includes a compact, executable session-first command sequence and a purpose-to-command capability map so an agent unfamiliar with xlflow can safely begin the normal development loop without first reconstructing it from command documentation.

The normal loop is: orient, analyze structural impact when relevant, edit source artifacts, validate, push or import, run focused verification, inspect workbook results when relevant, perform broader verification, then save and stop the session.

The installed Skill loads specialized guidance only when the task requires it:

| Task | Reference |
| ---- | --------- |
| Tests and test failures | `references/testing.md` |
| Formula-driven behavior | `references/formulas.md` |
| UserForm design and inspection | `references/forms.md` |
| Interactive dialogs and unattended UI | `references/xlflow-ui.md` |
| Runtime diagnostics | `references/debugging.md` |
| Recovery-required state | `references/recovery.md` |
| Call graph, impact, and dependencies | `references/code-analysis.md` |

This structure keeps detailed workflows available without turning `SKILL.md` into a CLI reference. In particular, recovery remains fail-closed: when xlflow reports that recovery is required, stop normal workbook operations and follow the recovery reference.

Supported provider targets:

| Agent target | Install path            |
| ------------ | ----------------------- |
| `agents`     | `.agents/skills/xlflow` |
| `codex`      | `.codex/skills/xlflow`  |
| `claude`     | `.claude/skills/xlflow` |
| `cursor`     | `.cursor/skills/xlflow` |
| `gemini`     | `.gemini/skills/xlflow` |

Use `--force` only when you intentionally want to overwrite an existing installed skill.
