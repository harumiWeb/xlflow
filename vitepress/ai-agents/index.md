# AI Agents

xlflow gives coding agents a stable interface for Excel VBA projects: source files, JSON output, explicit exit codes, diagnostics, sessions, and visual inspection commands.

Every command can return a stable JSON envelope, so coding agents can parse results without scraping terminal text. See [JSON Output](../reference/json-output) and [Error Codes](../reference/error-codes).

Before editing, confirm workbook and source state:

```bash
xlflow status --json
```

Recommended loop:

```bash
xlflow doctor --json
xlflow status --json
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

When the agent runs inside WSL, keep the repository under `/mnt/<drive>/...` and install xlflow on both WSL and Windows. The same commands transparently delegate Excel operations to Windows while `lint`, `fmt`, and `analyze` remain in WSL. Prefer the session-first sequence:

```bash
xlflow doctor --json
xlflow session start --json
xlflow push --fast --session --no-save --json
xlflow run Main.Run --diagnostic --session --json
xlflow inspect workbook --session --json
xlflow save --session --json
xlflow session stop --json
```

If a run fails, check state and recover with a short loop:

```bash
xlflow status --json
xlflow lint --json
xlflow analyze --json
xlflow run Main.Run --diagnostic --session --json
xlflow inspect workbook --json
xlflow pull --session --json
```

Use `run --diagnostic` when you need structured compile or runtime failure details instead of a blind Excel/VBE dialog. Re-run `pull` after workbook-side investigation when you need the latest saved workbook state back in source.

Install the bundled skill when supported by your agent:

```bash
xlflow skill install --agent codex
```
