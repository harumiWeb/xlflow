# Reporting bugs

Before opening an issue, reproduce with the smallest disposable workbook. This protects confidential data and makes the failure much faster to diagnose. Capture:

```powershell
xlflow version --verbose --json
xlflow doctor --json
xlflow status --json
```

Include the exact command, exit code, sanitized JSON envelope, workbook format (`.xlsm`, `.xlam`, or `.xlsb`), managed/external session state, and whether Excel remained running. Say what you expected and what actually happened. Never attach confidential workbooks or secrets; provide a minimal reproduction instead.
