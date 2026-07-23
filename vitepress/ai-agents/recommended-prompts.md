# Recommended Prompts

Use prompts that name the desired workbook result and force the agent to respect source-of-truth and verification boundaries. A vague “fix the macro” prompt invites guessing; an observable expectation gives the agent a way to stop.

```text
Use xlflow for this Excel VBA project. Read xlflow.toml first. Prefer --json. Run xlflow status --json before editing; if source and workbook freshness are unclear, stop and report it rather than overwriting either side. After edits, run fmt/lint/analyze, push, test or run, and inspect or export-image when workbook output matters. Do not save a session until the inspected result matches the requested outcome.
```

For debugging:

```text
Do not guess the macro name. Run xlflow macros --json, use the discovered qualified_name, and use XlflowDebug.Log plus xlflow run --diagnostic --json if the runtime failure is unclear. If error.code is workbook_recovery_required, stop normal commands and follow the recovery instructions.
```
