# Architecture

xlflow is a Windows-first Go CLI using Cobra for command routing, `xlflow.toml` for project configuration, stable JSON envelopes for automation, and a Windows bridge layer that now prefers the `.NET` Excel bridge while retaining PowerShell as a legacy fallback.

Key decisions:

- Excel-dependent commands require Windows, Excel, the configured bridge provider, and VBIDE trust settings.
- Non-Excel source checks remain testable without Excel.
- GUI interactions are explicit boundaries, not hidden automation targets.
- VBE compile diagnostics are handled narrowly because they are editor diagnostics, not workbook business UI.

Source ADR: [`docs/adr/ADR-0001-agent-ready-vba-cli-architecture.md`](https://github.com/harumiWeb/xlflow/blob/main/docs/adr/ADR-0001-agent-ready-vba-cli-architecture.md)
