# .NET Excel Bridge

The .NET Excel bridge is the Windows automation provider for workbook-backed xlflow commands.

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

## Implemented Commands

### `doctor`

`xlflow doctor --bridge dotnet --json` runs environment diagnostics through the .NET bridge without launching PowerShell. The response includes a `diagnostics` object at the top level:

```json
{
  "status": "ok",
  "command": "doctor",
  "diagnostics": {
    "selected_bridge": "dotnet",
    "protocol_version": 1,
    "runtime": {
      "os": "Windows 11",
      "process_architecture": "X64",
      "dotnet_runtime": ".NET 8.0"
    },
    "excel": {
      "com_activation": true,
      "version": "16.0",
      "build": "12345",
      "vbide_access": true,
      "automation_security": 1,
      "trust_vba_access": null
    }
  }
}
```

When Excel COM activation fails, the bridge returns a structured error with `code`, `message`, `phase`, `source`, `number`, `h_result`, and `details` fields. The `h_result` field is a hex HRESULT string (e.g. `"0x80040154"`).

Fields:

- `selected_bridge` — always `"dotnet"` when the .NET bridge handles the command.
- `protocol_version` — bridge protocol version (currently `1`).
- `runtime.os` — `Environment.OSVersion` string.
- `runtime.process_architecture` — process architecture (e.g. `X64`).
- `runtime.dotnet_runtime` — `.NET` runtime description.
- `excel.com_activation` — `true` when `Excel.Application` was successfully created on the STA thread established by `[STAThread]` in `Program.cs`. This encompasses both COM object creation and the STA execution context required by Excel.
- `excel.version` — Excel application version.
- `excel.build` — Excel application build number.
- `excel.vbide_access` — `true` when the VBA project object model is accessible.
- `excel.automation_security` — observed `AutomationSecurity` value.
- `excel.trust_vba_access` — observed Trust access state; `null` when not determinable.
- `excel.error` — present only when a non-fatal diagnostic warning occurred.

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
