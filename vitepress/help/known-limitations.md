# Known limitations

These are product boundaries, not usually setup errors. If your situation does not match one of them, run `xlflow doctor --json` and consult [Troubleshooting](./troubleshooting) before concluding that xlflow is unsupported.

- Workbook-backed commands require Windows desktop Excel and VBIDE access.
- The VS Code extension is Windows-only.
- WSL Excel delegation requires a Windows-mounted project path.
- Session test isolation is intentionally limited; retries requiring a fresh workbook baseline run outside a live session.
- CodeLens excludes argument-bearing procedures, functions, and properties.
- Some UserForm Designer properties are best-effort or snapshot-only; consult the UserForm specification.
- Headless automation cannot safely complete arbitrary third-party dialogs; use explicit XlflowUI wrappers or interactive mode with a human.
