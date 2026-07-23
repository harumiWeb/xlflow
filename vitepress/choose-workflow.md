# Choose your workflow

xlflow has one project model but several good entry points. Choose the path that matches what you need to do first. You do not have to learn every command before choosing: each path explains the commands only when you need them.

| Goal                                           | Start here                                                   | You will finish with                                  |
| ---------------------------------------------- | ------------------------------------------------------------ | ----------------------------------------------------- |
| Completion, diagnostics, navigation, and tests | [Develop VBA in VS Code](./tutorials/vscode-development)     | A project managed from the VS Code editor and sidebar |
| Have an agent implement and verify VBA         | [Develop with an AI agent](./tutorials/ai-agent)             | A repeatable JSON/session/inspection loop             |
| Put an existing `.xlsm` under Git              | [Import an existing workbook](./tutorials/existing-workbook) | `src/`, `build/`, `xlflow.toml`, and a reviewed diff  |
| Script Excel from a terminal or CI             | [CLI automation](./guides/ci)                                | Headless commands with exit codes and JSON output     |
| Keep source tools in Linux/WSL                 | [Work from WSL](./tutorials/wsl)                             | WSL linting with Windows Excel delegation             |

All paths use the same state model: source files are edited, `push` imports them, a live session can run repeatedly, and `save --session` persists workbook changes. Read [Source, workbook, and session state](./concepts/workbook-session-source) before switching between Excel and VS Code.

## Before you start

- Windows with Microsoft Excel is required for workbook-backed commands.
- Enable **Trust access to the VBA project object model** in Excel Trust Center.
- Run `xlflow version` and `xlflow doctor --json` before investigating a project failure.

If a prerequisite fails, use [Troubleshooting](./help/troubleshooting) rather than retrying an unsafe workbook operation.

### Not sure which one applies?

Most people with an existing macro-enabled workbook should choose **Manage an existing workbook with Git**. Choose **VS Code** when the editor experience is the immediate goal; choose **AI agent** when a terminal-driven agent will make and verify the change. You can move between these paths later because they all use the same `src/`, workbook, and session model.
