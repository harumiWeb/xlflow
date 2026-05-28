# PowerShell Bridge Legacy Fallback

The PowerShell bridge is xlflow's current Excel automation implementation and remains the active runtime until .NET provider work is implemented command by command.

This document defines its planned role after the .NET bridge migration begins.

## Role

PowerShell remains a supported legacy fallback while .NET reaches command parity.

It should continue to support existing workbook-backed behavior, including:

- `doctor`
- `pull`
- `push`
- `run`
- `test`
- `trace`
- `session`
- `macros`
- `inspect`
- `form`
- `export-image`
- `process`

New Windows automation capabilities should be designed .NET-first unless they are required to keep legacy fallback behavior safe.

## Selection

Users can force the PowerShell bridge with:

```powershell
xlflow run Main.Run --bridge powershell --json
```

In `auto` mode, PowerShell can be used when:

- The .NET bridge executable is missing.
- The .NET bridge protocol is incompatible.
- The .NET bridge does not support the requested command.
- The project is still in a migration phase where PowerShell is the preferred default.

In explicit `--bridge dotnet` mode, xlflow must not fallback to PowerShell. That strictness is required so users can verify that PowerShell is not involved in locked-down environments.

## stdout Contract

The public xlflow stdout contract remains a single final human output or JSON envelope. PowerShell scripts must not emit stray stdout outside their structured JSON result when invoked by xlflow.

Progress, debug, and UI stream data belong on stderr or inside final structured fields.

## Maintenance Policy

During migration:

- Fix regressions and security/safety problems in PowerShell.
- Preserve existing command output and exit-code behavior.
- Avoid expanding large shared PowerShell helpers for new .NET-first features.
- Keep tests for legacy fallback and explicit `--bridge powershell` paths.

After .NET becomes the default bridge on Windows, PowerShell documentation should describe it as legacy fallback rather than the primary automation path.

## Known Limitations

- Corporate execution policy or endpoint controls may block PowerShell script execution.
- Some helper behavior depends on Windows PowerShell compatibility.
- Complex COM, VBIDE, Win32 dialog, and clipboard workflows are harder to maintain safely in PowerShell than in C#.
