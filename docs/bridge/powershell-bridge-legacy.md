# Deprecated PowerShell Bridge

The PowerShell bridge is deprecated in v0.15.0 and is planned for removal in v0.16.0.

## Role During v0.15.0

The `.NET` bridge is the supported Windows automation bridge for workbook-backed xlflow commands. Windows `auto` mode selects `.NET` and does not fall back to PowerShell.

PowerShell remains available only as an explicit compatibility opt-in during the v0.15.0 deprecation window:

```powershell
xlflow run Main.Run --bridge powershell --json
```

The same deprecated opt-in can be selected through `XLFLOW_EXCEL_BRIDGE=powershell` or `[excel].bridge = "powershell"`. Any PowerShell bridge selection emits a `powershell_bridge_deprecated` warning that states the v0.16.0 removal target.

## Maintenance Policy

During the deprecation window:

- no new PowerShell bridge features are added;
- no behavior changes are made except critical safety or regression fixes;
- no new compatibility guarantees are added;
- missing `.NET` bridge runtime, protocol, capability, transport, or workbook-open failures are reported directly instead of falling back to PowerShell.

Workflows that still require PowerShell should report the blocker so the missing behavior can be ported to the `.NET` bridge before v0.16.0.

## Known Limitations

- Corporate execution policy or endpoint controls may block PowerShell script execution.
- Some helper behavior depends on Windows PowerShell compatibility.
- Complex COM, VBIDE, Win32 dialog, and clipboard workflows are harder to maintain safely in PowerShell than in C#.
