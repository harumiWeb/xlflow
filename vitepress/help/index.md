# Help and troubleshooting

Use this section when xlflow is installed but a command, workbook, session, or editor feature is not behaving as expected. Start with the symptom—not with a risky retry.

## First three safe checks

```powershell
xlflow version --verbose --json
xlflow doctor --json
xlflow status --json
```

`version` confirms which CLI is running, `doctor` checks the Excel and VBIDE environment, and `status` reports the relationship between source files, the saved workbook, and a live session. These commands do not overwrite workbook VBA.

## Find the right kind of help

| If you are trying to…                                          | Start here                               |
| -------------------------------------------------------------- | ---------------------------------------- |
| Fix an installation, Excel, execution, session, or WSL failure | [Troubleshooting](./troubleshooting)     |
| Understand a common xlflow question                            | [FAQ](./faq)                             |
| Check whether a capability is currently supported              | [Known limitations](./known-limitations) |
| Report a reproducible problem safely                           | [Reporting bugs](./reporting-bugs)       |

When a command returns JSON with `"status": "error"`, use its `error.code` and the command's exit code to select the relevant troubleshooting section. If `error.code` is `workbook_recovery_required`, stop normal workbook commands and follow the recovery guidance; retrying with `--wait` cannot clear that state.
