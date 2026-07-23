# Install the VS Code extension

1. Install xlflow and the Excel bridge using [Installation](../installation).
2. Enable **Trust access to the VBA project object model** in Excel.
3. Install the `xlflow` extension from the VS Code Extensions view or the packaged release.
4. Open a folder containing `xlflow.toml`.
5. Run `xlflow: Check Environment` or `xlflow doctor --json`.

If the CLI is not on `PATH`, set:

```json
{ "xlflow.path": "C:\\Users\\me\\AppData\\Local\\xlflow\\xlflow.exe" }
```

Reload VS Code after changing the path. The extension is Windows-only because workbook-backed operations use Excel COM.
