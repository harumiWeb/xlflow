# Codex

Install the Codex skill into the repository:

```bash
xlflow skill install --agent codex
```

Codex should treat `src/` as the primary edit surface, prefer JSON output, and use `xlflow session` for repeated workbook-backed loops. For release-grade workbook behavior changes, run Windows + Excel COM E2E in a disposable workspace.
