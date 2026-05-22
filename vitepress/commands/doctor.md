# xlflow doctor

Diagnose Excel, COM, PowerShell, VBIDE access, workbook access, and GUI-boundary prerequisites.

## Usage

```bash
xlflow doctor
```

## Options and Arguments

| Option / argument | Description                                           | Default |
| ----------------- | ----------------------------------------------------- | ------- |
| `--json`          | Return structured diagnostics for agents and CI logs. | false   |

## Examples

```bash
xlflow doctor
```

## Notes

::: tip
Run `doctor` before debugging workbook behavior on a new Windows or Excel installation.
:::

::: warning
VBIDE access must be enabled in Excel Trust Center before xlflow can import, export, or inspect VBA components.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "doctor",
  "checks": [
    { "name": "excel_com", "status": "ok" },
    { "name": "vbide_access", "status": "ok" }
  ]
}
```

## Related

- [check](./check)
- [troubleshooting](../reference/troubleshooting)
- [environment variables](../reference/environment-variables)
