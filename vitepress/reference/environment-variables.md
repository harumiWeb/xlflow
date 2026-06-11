# Environment Variables

| Variable                 | Purpose                                                                                 |
| ------------------------ | --------------------------------------------------------------------------------------- |
| `XLFLOW_NO_UPDATE_CHECK` | Set to `1` to disable the interactive release update check during `new` and `init`.     |
| `XLFLOW_WINDOWS_EXE`     | Under WSL, override the Windows `xlflow.exe` used for delegated Excel-related commands. |

Workbook-backed commands also reflect the local Excel, COM, PowerShell, and VBIDE environment. Run `xlflow doctor --json` before debugging source when environment setup is uncertain.
