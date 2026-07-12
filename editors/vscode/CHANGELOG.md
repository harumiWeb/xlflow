# Changelog

## Unreleased

- Added VBA annotation completion for xlflow `@ExpectedError(...)`, `@Skip(...)`, and `@Todo(...)` test metadata when typing `@` in apostrophe comments.
- Added VS Code editor support for xlflow VBA documentation comments, including Quick Fix snippet generation from `'''`, Rubberduck `@Description` annotation completions in comments, doc-comment continuation on Enter, and highlighting for `'''` comments and Rubberduck description annotations.
- Updated Test Explorer to use xlflow's stable qualified VBA test IDs and run focused tests with `xlflow test --filter Module.TestName`.

## v0.5.0

- Added editor diagnostics for high-signal analyze warnings, including `VBA201`, `VBA204`, `VBA208`, `VBA209`, and `VBA212`.
- Suppressed Format Document failure notifications for incomplete or syntactically invalid VBA buffers; formatting now leaves the document unchanged instead.
- Added a VS Code Formulas sidebar view for discovering generated formula snapshots, opening `names.jsonl` and per-sheet region JSONL files, and running `xlflow formulas pull`.

## v0.4.0

- Added setup warnings for disabled Excel VBA object model access and for `.bas`, `.cls`, or `.frm` files that are not opened with xlflow's `vba` language mode.
- Added a Project view session action to connect xlflow to an already-open workbook with `xlflow session attach`.
- Open the configured workbook from the Project view in the associated desktop app instead of opening the binary workbook inside VS Code.

## v0.3.0

- Added Quick Fix actions for xlflow diagnostics to insert `xlflow:disable-next-line` or `xlflow:disable-line` suppression comments.
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
