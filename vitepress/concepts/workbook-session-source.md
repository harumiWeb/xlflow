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

To check the current state of all three layers at once, use `status`:

```bash
xlflow status --json
```

`status` shows whether source files are newer than the saved workbook, whether a live session is active and dirty, and which path is the current source of truth.
