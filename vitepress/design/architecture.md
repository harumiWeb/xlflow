# Architecture

xlflow is a Windows-first Go CLI using Cobra for command routing, `xlflow.toml` for project configuration, stable JSON envelopes for automation, and a Windows bridge layer that uses the `.NET` Excel bridge in `auto` mode. The legacy PowerShell bridge was removed in v0.16.0. Under WSL, source-only commands remain local while Excel-related commands delegate to Windows `xlflow.exe`.

Key decisions:

- Excel-dependent commands require Windows, Excel, the configured bridge provider, and VBIDE trust settings.
- WSL projects that use Excel delegation must live under a Windows-mounted path such as `/mnt/c/...`.
- Non-Excel source checks remain testable without Excel.
- GUI interactions are explicit boundaries, not hidden automation targets.
- VBE compile diagnostics are handled narrowly because they are editor diagnostics, not workbook business UI.

Source ADR: [`docs/adr/ADR-0001-agent-ready-vba-cli-architecture.md`](https://github.com/harumiWeb/xlflow/blob/main/docs/adr/ADR-0001-agent-ready-vba-cli-architecture.md)

WSL delegation ADR: [`docs/adr/ADR-0011-wsl-windows-cli-delegation.md`](https://github.com/harumiWeb/xlflow/blob/main/docs/adr/ADR-0011-wsl-windows-cli-delegation.md)
