# Develop with an AI agent

This tutorial gives an agent a bounded task: add a `Main.CalculateTotal` procedure that writes the total of `Input!B2:B10` to `Result!B2`, then prove the result.

## Prompt

```text
You are working in an xlflow project. Read xlflow.toml, inspect src/, and run
xlflow status --json before editing. Add Main.CalculateTotal and a TestCalculateTotal
test. Use xlflow fmt, lint, and analyze before touching Excel. Start one session,
push with --session --no-save, run the test and macro with --diagnostic, inspect
Result!B2, export an image of the result, then save the session and show git diff.
If a command fails, parse error.code and fix the source; never wait for a GUI dialog.
```

## Expected loop

```bash
xlflow status --json
xlflow lint --json
xlflow analyze --json
xlflow test --json
xlflow session start --json
xlflow push --fast --session --no-save --json
xlflow test --session --json
xlflow run Main.CalculateTotal --diagnostic --session --json
xlflow inspect range --sheet Result --address B2 --session --json
xlflow export-image --sheet Result --range A1:C5 --session --json
xlflow save --session --json
xlflow session stop --json
git diff -- src
```

The first test should fail if the procedure is missing. The agent uses the structured failure, edits source, and reruns the same session. `--no-save` keeps the experiment reversible; `save --session` is the explicit persistence boundary. If execution returns `workbook_recovery_required`, stop the loop and use [recovery](../commands/recovery), not `--wait`.
