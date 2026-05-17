# AI Agents

xlflow gives coding agents a stable interface for Excel VBA projects: source files, JSON output, explicit exit codes, diagnostics, sessions, and visual inspection commands.

Recommended loop:

```bash
xlflow doctor --json
xlflow session start --json
xlflow pull --session --json
xlflow push --fast --session --no-save --json
xlflow lint --json
xlflow analyze --json
xlflow test --session --json
xlflow macros --session --json
xlflow run Main.Run --headless --session --json
xlflow save --session --json
```

Install the bundled skill when supported by your agent:

```bash
xlflow skill install --agent codex
```
