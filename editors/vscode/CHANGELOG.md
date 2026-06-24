# Changelog

## 0.1.0

- Added the initial production VS Code extension for xlflow.
- Registered `.bas`, `.cls`, and `.frm` as VBA files.
- Added a thin `vscode-languageclient` client for `xlflow lsp --stdio`.
- Added command palette actions for restarting the language server, checking the environment, pulling workbook sources, pushing sources, and running the configured macro.
- Added command palette actions for creating projects, initializing projects from existing workbooks, and starting, checking, or stopping workbook sessions.
- Added a command palette action for saving the managed workbook with `xlflow save`.
- Added a command palette action for running workbook tests with `xlflow test` and showing results in the output channel.
- Added VS Code Test Explorer integration backed by `xlflow test list --json` and `xlflow test --json`.
- Added command palette actions for linting the workspace and formatting the active document or full project.
- Added a Status Bar session indicator with QuickPick actions for starting, stopping, restarting, inspecting, and diagnosing xlflow sessions.
