# Workbook, Session, Source

xlflow makes three states explicit:

```text
source files
   | push
   v
Excel workbook file
   | session start / run / save
   v
live Excel session
   | pull
   v
source files
```

Use `pull` when the workbook may contain newer VBA than the source tree. Use `push` after editing source files. Use `session` for repeated workbook-backed commands without reopening Excel.

When a session-backed command leaves the live workbook newer than disk, xlflow reports save-required metadata. Persist that state with:

```bash
xlflow save --session --json
```
