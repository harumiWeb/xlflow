# xlflow doctor

Diagnose Excel, COM, `.NET` bridge, VBIDE access, systemprofile Desktop readiness, optional workbook access, and GUI-boundary prerequisites.

## Usage

```bash
xlflow doctor [--bridge <auto|dotnet>] [--workbook]
```

## Options and Arguments

| Option / argument     | Description                                                    | Default |
| --------------------- | -------------------------------------------------------------- | ------- |
| `--json`              | Return structured diagnostics for agents and CI logs.          | false   |
| `--bridge <provider>` | Select the Excel bridge provider (`auto`, `dotnet`).           | `auto`  |
| `--workbook`          | Open and close the configured workbook as part of diagnostics. | false   |

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
By default, `doctor` performs Excel COM, VBIDE, temporary-workbook VBProject, and systemprofile Desktop checks without opening the configured workbook. Use `--workbook` when you need to prove that the configured workbook can be opened and its VBProject can be accessed by the selected bridge.
:::

::: tip
`doctor` can run outside an xlflow project. When `xlflow.toml` is missing, it still checks the project-independent Excel, bridge, VBIDE, and TypeLib environment, reports `diagnostics.project_config.found: false`, and prints hints for `xlflow new` or `xlflow init <workbook>`. In that state, `--workbook` is skipped because there is no configured workbook path.
:::

::: tip
`doctor` also checks the generated TypeLib database used by the LSP. If it has not been initialized, or if it is stale, `doctor` keeps the environment check successful but prints a warning and a hint to run `xlflow type db init` or `xlflow type db refresh --library all`.
:::

::: warning
VBIDE access must be enabled in Excel Trust Center before xlflow can import, export, inspect, or run VBA components. When `doctor` can determine that access is disabled, it reports warning `vba_object_model_access_disabled` with a hint to enable **Trust access to the VBA project object model**.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus the `diagnostics` object.

On Windows, `doctor` uses the `.NET` bridge in `auto` mode. The nested `diagnostics` object always reports the requested bridge mode, the selected provider, `fallback=false`, and `legacy=false`.

When no `xlflow.toml` exists in the current workspace, successful output also includes `diagnostics.project_config.found: false` plus warnings and hints. Project-dependent checks, including workbook opening, are skipped.

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

The removed `--bridge powershell` value fails before Excel starts. Use `auto` or `dotnet`.

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

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow doctor` when the task matches the command description above. For a goal-oriented workflow, start with the [How-to guides](../guides/) and return here for exact options.

## Prerequisites

Check the project configuration and run `xlflow doctor --json` before workbook-backed operations. Source-only commands can run without Excel; commands that read or mutate a workbook require Windows Excel and VBIDE access.

## What this command reads and changes

The command reads the inputs and configuration described in its syntax and examples. Treat source files, the saved workbook, and a live session as separate states; add `--session` when the live workbook is authoritative. Any mutation is reversible only when a backup or explicit session save boundary exists.

## Effect on source-of-truth state

Use `xlflow status --json` before and after the command. A source edit normally requires `push`; a workbook edit normally requires `pull`; a dirty live session requires `save --session` or an intentional discard.

## Common workflows

Combine this command with the relevant [source/workbook/session workflow](../concepts/workbook-session-source), and use `--json` in scripts and agent loops.

## Common failures

Read the structured `error.code`, exit code, and recovery metadata instead of scraping terminal text. The [symptom-oriented troubleshooting guide](../help/troubleshooting) maps installation, execution, session, VS Code, and WSL failures to recovery steps.
