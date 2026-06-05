# PowerShell Bridge Legacy Fallback

The PowerShell bridge is xlflow's supported legacy Excel automation fallback on Windows.

## Role

PowerShell remains usable, but it is no longer the default Windows bridge. Windows `auto` mode now prefers `.NET`, and PowerShell is used only when selected explicitly or when `auto` must fall back.

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

New feature work should not expand the PowerShell bridge unless it is required for fallback parity, safety, or bug fixes.

## Known Limitations

- Corporate execution policy or endpoint controls may block PowerShell script execution.
- Some helper behavior depends on Windows PowerShell compatibility.
- Complex COM, VBIDE, Win32 dialog, and clipboard workflows are harder to maintain safely in PowerShell than in C#.
