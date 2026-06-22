# xlflow doctor

Diagnose Excel, COM, `.NET` bridge, VBIDE access, systemprofile Desktop readiness, optional workbook access, and GUI-boundary prerequisites.

## Usage

```bash
xlflow doctor [--bridge <auto|dotnet>] [--workbook]
```

## Options and Arguments

| Option / argument     | Description                                                                                                         | Default |
| --------------------- | ------------------------------------------------------------------------------------------------------------------- | ------- |
| `--json`              | Return structured diagnostics for agents and CI logs.                                                               | false   |
| `--bridge <provider>` | Select the Excel bridge provider (`auto`, `dotnet`). Deprecated `powershell` remains explicit opt-in until v0.16.0. | `auto`  |
| `--workbook`          | Open and close the configured workbook as part of diagnostics.                                                      | false   |

## Examples

```bash
xlflow doctor
xlflow doctor --bridge dotnet --json
xlflow doctor --workbook
```

## Notes

::: tip
Run `doctor` before debugging workbook behavior on a new Windows or Excel installation.
:::

::: tip
By default, `doctor` performs lightweight Excel COM, VBIDE, and systemprofile Desktop checks without opening the configured workbook. Use `--workbook` when you need to prove that the configured workbook can be opened by the selected bridge.
:::

::: warning
VBIDE access must be enabled in Excel Trust Center before xlflow can import, export, or inspect VBA components.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus the `diagnostics` object.

On Windows, `doctor` uses the `.NET` bridge in `auto` mode. The nested `diagnostics` object always reports the requested bridge mode, the selected provider, and `fallback=false` because automatic PowerShell fallback is disabled in v0.15.0.

Under WSL, `doctor` invokes Windows xlflow and augments the Windows result with:

- `diagnostics.host`: Linux/WSL detection, distro, and WSL xlflow version.
- `diagnostics.windows`: Windows xlflow path/version plus bridge and Excel availability.
- `diagnostics.path_translation`: WSL and Windows project paths and support status.

A version mismatch between WSL and Windows xlflow is reported as a warning but does not block delegation.

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
      "trust_vba_access": null,
      "systemprofile_desktop": {
        "system32": {
          "path": "C:\\Windows\\System32\\config\\systemprofile\\Desktop",
          "status": "exists"
        },
        "syswow64": {
          "path": "C:\\Windows\\SysWOW64\\config\\systemprofile\\Desktop",
          "status": "exists"
        },
        "ok": true
      }
    }
  }
}
```

With `--workbook`, successful output also includes the configured workbook path and `diagnostics.workbook_openable: true`.

WSL output includes the same Windows diagnostics plus the delegation boundary:

```json
{
  "status": "ok",
  "command": "doctor",
  "diagnostics": {
    "host": {
      "os": "linux",
      "is_wsl": true,
      "distro": "Ubuntu-24.04"
    },
    "windows": {
      "xlflow_found": true,
      "xlflow_path": "C:\\Users\\me\\AppData\\Local\\xlflow\\xlflow.exe",
      "xlflow_version": "0.13.0",
      "bridge_found": true,
      "excel_available": true
    },
    "path_translation": {
      "supported": true,
      "wsl_path": "/mnt/c/dev/project",
      "windows_path": "C:\\dev\\project"
    }
  }
}
```

When deprecated `--bridge powershell` is selected explicitly, the top-level `bridge` metadata uses the PowerShell provider shape, `diagnostics.legacy=true`, and the response includes a `powershell_bridge_deprecated` warning:

```json
{
  "status": "ok",
  "command": "doctor",
  "diagnostics": {
    "requested_bridge": "powershell",
    "selected_bridge": "powershell",
    "fallback": false,
    "legacy": true
  },
  "bridge": {
    "host": "pwsh.exe",
    "edition": "Core",
    "version": "7.5.1"
  },
  "warnings": [
    {
      "code": "powershell_bridge_deprecated",
      "message": "The PowerShell bridge is deprecated in v0.15.0 and will be removed in v0.16.0. Use the .NET bridge.",
      "removal_version": "v0.16.0"
    }
  ]
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

When the Windows systemprofile Desktop directories are missing, `doctor` fails with an actionable environment error. These directories are required for Excel COM workbook automation in non-interactive sessions such as SSH, services, and CI.

```json
{
  "status": "failed",
  "command": "doctor",
  "error": {
    "code": "systemprofile_desktop_missing",
    "message": "systemprofile Desktop directories are missing.\nCreate both directories:\n- C:\\Windows\\System32\\config\\systemprofile\\Desktop\n- C:\\Windows\\SysWOW64\\config\\systemprofile\\Desktop\n\nThis is required for Excel COM automation in non-interactive sessions such as SSH, services, or CI.",
    "phase": "doctor",
    "source": "xlflow-excel-bridge"
  },
  "diagnostics": {
    "excel": {
      "com_activation": true,
      "systemprofile_desktop": {
        "system32": {
          "path": "C:\\Windows\\System32\\config\\systemprofile\\Desktop",
          "status": "missing"
        },
        "syswow64": {
          "path": "C:\\Windows\\SysWOW64\\config\\systemprofile\\Desktop",
          "status": "missing"
        },
        "ok": false,
        "missing": true
      }
    }
  }
}
```

If a systemprofile Desktop path exists but cannot be inspected by the current user, `status` is `access_denied`. That condition is reported as a warning in human output; it does not by itself make `doctor` fail.

## Related

- [check](./check)
- [troubleshooting](../reference/troubleshooting)
- [environment variables](../reference/environment-variables)
