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

When `--bridge dotnet` is selected explicitly, the top-level `bridge` metadata identifies the .NET bridge process and the nested `diagnostics` object contains the .NET-specific runtime and Excel probe results shown below.

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

When `--bridge powershell` is selected, or when `--bridge auto` resolves to PowerShell, the top-level `bridge` metadata uses the PowerShell provider shape instead:

```json
{
  "status": "ok",
  "command": "doctor",
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
