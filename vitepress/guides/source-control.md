# Manage VBA source with Git

Treat `src/` and `xlflow.toml` as the reviewable description of your VBA project. Treat `build/` as the workbook artifact and `.xlflow/` as generated working state. This lets a code review explain a VBA change without asking reviewers to open an Excel file.

```bash
xlflow pull --json          # workbook -> source
git diff -- src
xlflow fmt --check
xlflow lint --json
xlflow push --json          # source -> workbook
git diff -- src xlflow.toml
```

`pull` is how an existing workbook enters Git; it does not send source back into Excel. After reviewing the exported baseline, commit it before a risky push so Git and xlflow's workbook backup both provide recovery options. If you do not know which side is newer, run `xlflow status --json`, make a backup, and pull only after confirming the workbook is authoritative.
