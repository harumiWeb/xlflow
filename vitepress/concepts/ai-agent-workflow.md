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
- Replace raw `MsgBox` / `InputBox` with `XlflowUI` and pass `--msgbox` / `--inputbox` during unattended `run` or `test` flows.
- Use `inspect` and `export-image` to verify workbook output.
- Use `XlflowDebug.Log` plus `xlflow run --json` or `xlflow test --json` when runtime failures are unclear.
- Use `backup list` and `rollback` when a `push` or workbook mutation leaves the workbook in a bad state.
- Avoid headless workflows only for truly human-operated raw dialogs or UserForms; simple confirmation and scalar-input flows should move to `XlflowUI` first.
