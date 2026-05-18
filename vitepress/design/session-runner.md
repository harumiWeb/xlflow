# Session Runner

Sessions keep Excel and the configured workbook open across repeated commands. This makes AI-agent development loops faster while keeping save-required state explicit.

Typical loop:

```bash
xlflow session start --json
xlflow push --fast --session --no-save --json
xlflow run Main.Run --headless --session --json
xlflow save --session --json
xlflow session stop --json
```

Related ADR: [`ADR-0004`](https://github.com/harumiWeb/xlflow/blob/main/docs/adr/ADR-0004-explicit-excel-session-mode.md)
