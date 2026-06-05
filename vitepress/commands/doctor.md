# xlflow doctor

Diagnose Excel, COM, PowerShell, VBIDE access, workbook access, and GUI-boundary prerequisites.

## Usage

```bash
xlflow doctor [--bridge <auto|powershell|dotnet>]
```

## Options and Arguments

| Option / argument     | Description                                                        | Default |
| --------------------- | ------------------------------------------------------------------ | ------- |
| `--json`              | Return structured diagnostics for agents and CI logs.              | false   |
| `--bridge <provider>` | Select the Excel bridge provider (`auto`, `powershell`, `dotnet`). | `auto`  |

## Examples

```bash
xlflow doctor
xlflow doctor --bridge dotnet --json
```

## Notes

::: tip
Run `doctor` before debugging workbook behavior on a new Windows or Excel installation.
:::

::: warning
VBIDE access must be enabled in Excel Trust Center before xlflow can import, export, or inspect VBA components.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus the `diagnostics` object.

On Windows, `doctor` prefers the `.NET` bridge in `auto` mode. The nested `diagnostics` object always reports the requested bridge mode, the selected provider, and whether `auto` fell back to the legacy PowerShell bridge.

```json
{
  "status": "ok",
  "command": "doctor",
  "diagnostics": {
    "requested_bridge": "auto",
    "selected_bridge": "dotnet",
    "fallback": false,
    "legacy": false,
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

When `--bridge powershell` is selected, or when `--bridge auto` falls back to PowerShell, the top-level `bridge` metadata uses the PowerShell provider shape and `diagnostics.legacy=true`:

```json
{
  "status": "ok",
  "command": "doctor",
  "diagnostics": {
    "requested_bridge": "auto",
    "selected_bridge": "powershell",
    "fallback": true,
    "legacy": true
  },
  "bridge": {
    "host": "pwsh.exe",
    "edition": "Core",
    "version": "7.5.1"
  }
}
```

The provider-specific top-level `bridge` metadata identifies the bridge process. It is separate from `diagnostics`, which describes the probed Excel/runtime environment for the selected provider.

When Excel COM activation fails with `--bridge dotnet`, the response uses the standard error envelope:

```json
{
  "status": "failed",
  "command": "doctor",
  "error": {
    "code": "excel_com_failure",
    "message": "Excel COM activation failed",
    "phase": "doctor",
    "source": "xlflow-excel-bridge",
    "number": -2146959354,
    "h_result": "0x80040154"
  }
}
```

## Related

- [check](./check)
- [troubleshooting](../reference/troubleshooting)
- [environment variables](../reference/environment-variables)
