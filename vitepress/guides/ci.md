# Use xlflow in CI or automation

Install the CLI and bridge on the Windows runner, enable VBIDE access, and run source-only checks before workbook-backed checks.

```powershell
xlflow version --json
xlflow doctor --json
xlflow fmt --check --json
xlflow lint --json
xlflow analyze --json
xlflow test --json
```

Parse the JSON envelope and use exit codes rather than scraping human output. Never run parallel workbook mutations against the same file. For release-grade behavior, use a disposable workbook and the session-first workflow described in the repository E2E guidance.
