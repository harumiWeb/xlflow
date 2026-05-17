# AI Agent Workflow

xlflow is designed so agents can work through stable commands instead of the Excel UI.

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
xlflow session stop --json
```

Rules for agents:

- Prefer `--json`.
- Run `doctor` before changing source for environment failures.
- Use `macros` before guessing a `run` target.
- Use `inspect` and `export-image` to verify workbook output.
- Use `trace` when runtime failures are unclear.
- Avoid headless workflows for macros that intentionally require dialogs or UserForms.
