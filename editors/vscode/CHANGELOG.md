# Changelog

## Unreleased

- Offer a Push & Run choice before VS Code macro/procedure runs when source files are newer than the workbook.
- Run macros from VS Code with `xlflow run --interactive` so user-triggered macro execution uses interactive mode explicitly.
- Run user-triggered macro and procedure commands in the VS Code terminal instead of the Output panel.
- Keep the xlflow Output panel in the background for routine commands; diagnostics still open it explicitly, and failures offer an Open Output action.

## v0.2.0

- Added a visible pulse icon for the `Run Doctor` setup action in the VS Code sidebar.
- Renamed the Japanese VS Code setup action for existing workbook initialization to clarify that it converts a workbook into an xlflow project instead of wiping the workbook.
- Added xlflow CLI update notifications, a Project view update indicator, and a command palette action to check for updates manually.
- Added sidebar save actions so the Project view toolbar and `Save required` indicator can run `xlflow save` when a session workbook needs saving.

## 0.1.0

- Added CodeLens commands for running no-argument VBA procedures and tests through `xlflow.runProcedure` and `xlflow.runTestProcedure`.
- Added CodeLens and save-before-run settings forwarded to `xlflow lsp --stdio` at startup.
- Added command palette actions for `xlflow skill install` and `xlflow module install`.
- Added Modules TreeView context menu actions for opening, renaming, deleting, revealing, copying, and running module/procedure items.
- Added UserForms TreeView context menu actions for opening, renaming, deleting, revealing, and copying UserForm source artifacts.
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
- Tightened v0.1 review feedback by wiring LSP trace settings, gating startup test discovery, documenting Node 22 development requirements, and hardening xlflow process handling.
