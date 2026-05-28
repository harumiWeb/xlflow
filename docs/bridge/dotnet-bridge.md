# .NET Excel Bridge

The .NET Excel bridge is the planned Windows automation provider for workbook-backed xlflow commands.

It is introduced by ADR-0008 and uses the provider/fallback contract from ADR-0009.

## Responsibilities

The .NET bridge owns Windows automation concerns that are difficult to keep reliable in PowerShell:

- Excel COM automation.
- VBIDE automation.
- Workbook open/save/close and session attachment.
- Macro execution.
- UserForm import/export.
- Win32 and UI Automation dialog watching.
- Compile/runtime error capture.
- Process/window ownership checks.
- Clipboard and image export fallback behavior.

The Go CLI remains responsible for command parsing, config loading, source-tree decisions, static checks, public envelope mapping, and provider selection.

## Process Model

The bridge is a separate executable named `xlflow-excel-bridge.exe`.

The Go CLI starts it for command-scoped work and communicates through stdin/stdout JSON. stdout is reserved for protocol JSON only. Diagnostic logs go into response `logs` or stderr.

## COM Threading Rule

Excel COM operations must run in an STA context.

Implementation rules:

- Entry points that touch Excel COM must be `[STAThread]` or dispatch to a dedicated STA thread.
- Do not resume Excel COM work on arbitrary ThreadPool threads after `await`.
- Keep initial COM command handlers synchronous unless a dedicated STA dispatcher exists.
- Release COM references deliberately and keep workbook/session ownership explicit.

This rule is part of the public implementation contract because violating it can produce intermittent Excel/VBIDE failures that unit tests may not catch.

## Capabilities

The bridge advertises supported command keys through:

```powershell
xlflow-excel-bridge.exe --capabilities-json
```

The Go resolver must check capabilities before selecting .NET in `auto` mode.

## Selection

Users will be able to select the bridge through:

```powershell
xlflow doctor --bridge dotnet --json
xlflow run Main.Run --bridge dotnet --json
```

When `--bridge dotnet` is selected explicitly, xlflow must not fallback to PowerShell. Missing bridge executable, version mismatch, protocol mismatch, or unsupported command must be reported as structured errors.

## Initial Migration Order

The planned migration order is:

1. `doctor`
2. `inspect` and `process`
3. `pull` and `push`
4. `macros` and `run`
5. dialog watcher
6. `test`, `trace`, and runtime injection
7. `form` and `export-image`
8. release packaging
9. default bridge switch

`run` is intentionally not first because it combines macro invocation, runtime injection, compile checks, dialog capture, timeout behavior, and session handling.

## Release Notes

The .NET bridge avoids PowerShell execution policy, but it does not bypass all corporate controls. AppLocker, WDAC, Defender, EDR, or code-signing rules may still block an unsigned or unapproved executable. Release packaging and signing decisions must be documented before making .NET the default Windows bridge.
