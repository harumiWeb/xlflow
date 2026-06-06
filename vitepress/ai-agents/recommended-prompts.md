# Recommended Prompts

Use prompts that force the agent to respect source-of-truth and verification boundaries.

```text
Use xlflow for this Excel VBA project. Read xlflow.toml first. Prefer --json. If source and workbook freshness are unclear, run xlflow pull --json before editing. After edits, run lint/analyze, push, test or run, and inspect or export-image when workbook output matters.
```

For debugging:

```text
Do not guess the macro name. Run xlflow macros --json, use the discovered qualified_name, and use XlflowDebug.Log plus xlflow run --json if the runtime failure is unclear.
```
