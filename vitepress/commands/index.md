# Commands

xlflow exposes a Cobra-based CLI. Every command accepts the global `--json` flag for machine-readable output.

```bash
xlflow [command] --json
```

Use command pages for workflow guidance and the canonical CLI contract in [JSON Output](../reference/json-output).

| Command                        | Purpose                                                                          |
| ------------------------------ | -------------------------------------------------------------------------------- |
| [new](./new)                   | Create a new xlflow project and macro-enabled workbook.                          |
| [init](./init)                 | Initialize an xlflow project from an existing workbook.                          |
| [doctor](./doctor)             | Diagnose Excel, COM, PowerShell, VBIDE access, and source GUI boundaries.        |
| [attach](./attach)             | Validate that the active Excel workbook matches the configured workbook.         |
| [backup](./backup)             | List rollback-capable workbook backups for the configured workbook.              |
| [list](./list)                 | List workbook resources. The public subcommand is `list forms`.                  |
| [form](./form)                 | Manage UserForms through snapshot, build, and image export workflows.            |
| [pull](./pull)                 | Export workbook VBA components into configured source directories.               |
| [push](./push)                 | Import edited source files back into the configured workbook.                    |
| [rollback](./rollback)         | Restore the configured workbook from an xlflow-managed backup.                   |
| [status](./status)             | Show project, source, workbook, and session state in one read-only command.      |
| [session](./session)           | Keep Excel and the configured workbook open across repeated commands.            |
| [save](./save)                 | Save the workbook held by the managed xlflow session.                            |
| [runner](./runner)             | Manage the persistent xlflow runner marker module.                               |
| [trace](./trace)               | Manage VBA trace logging support and trace log cleanup.                          |
| [run](./run)                   | Run a workbook macro from the CLI.                                               |
| [export-image](./export-image) | Export a worksheet range as a PNG image through Excel COM.                       |
| [edit](./edit)                 | Mutate a live session workbook for setup and visual tuning.                      |
| [macros](./macros)             | Discover runnable public workbook macro entrypoints without executing user code. |
| [ui](./ui)                     | Manage xlflow-owned worksheet buttons.                                           |
| [test](./test)                 | Discover and run workbook VBA test procedures.                                   |
| [diff](./diff)                 | Compare workbook files and optionally exported VBA source trees.                 |
| [inspect](./inspect)           | Inspect saved workbook state or UserForm state.                                  |
| [inspect-gui](./inspect-gui)   | Scan source for automation-hostile GUI boundaries without opening Excel.         |
| [lint](./lint)                 | Lint VBA source files for agent-hostile and compile-dialog-prone patterns.       |
| [fmt](./fmt)                   | Format VBA source files with a conservative, non-destructive formatter.          |
| [analyze](./analyze)           | Analyze VBA source for runtime-risk patterns without Excel COM.                  |
| [check](./check)               | Run lint, analyze, and doctor as a combined preflight.                           |
| [module](./module)             | Install bundled xlflow helper modules into an existing project.                  |
| [completion](./completion)     | Generate shell completion scripts through Cobra.                                 |
| [process](./process)           | List and manage local Excel processes.                                           |
| [skill](./skill)               | Install the bundled xlflow skill for AI agent tools.                             |
| [version](./version)           | Show xlflow build metadata.                                                      |
