# xlflow macros

Discover runnable workbook macro entrypoints without executing user code.

## Usage

```bash
xlflow macros [--session] [--runnable]
```

## Options and Arguments

| Option / argument | Description                                         | Default |
| ----------------- | --------------------------------------------------- | ------- |
| `--session`       | Read macro metadata from the managed live workbook. | false   |
| `--runnable`      | Show only macros that are directly runnable.        | false   |
| `--json`          | Return full macro metadata as structured JSON.      | false   |

## Examples

```bash
xlflow macros
xlflow macros --session --json
xlflow macros --runnable --json
```

## Notes

Macro procedure names may use Unicode VBA identifiers; discovered `name` and `qualified_name` values preserve those names.

::: tip
Run `macros --json` before `run` so agents can choose an exact entrypoint. Use `--runnable` to filter out non-runnable procedures (those with parameters, event handlers, or on unsupported component types). The injected runner can invoke no-argument private, public, and friend procedures in standard, class, and workbook document modules.
:::

> [!IMPORTANT]
> This command inspects workbook state; it does not execute discovered macros.

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "macros",
  "default_entry": "Main.GenerateQRCode",
  "macros": [
    {
      "module": "Main",
      "name": "GenerateQRCode",
      "qualified_name": "Main.GenerateQRCode",
      "kind": "sub",
      "args": [],
      "line": 2,
      "component_type": "standard_module",
      "visibility": "Public",
      "has_parameters": false,
      "runnable": true,
      "reason_not_runnable": null,
      "run_command": "xlflow run Main.GenerateQRCode --session --json"
    },
    {
      "module": "Sheet1",
      "name": "Worksheet_Change",
      "qualified_name": "Sheet1.Worksheet_Change",
      "kind": "sub",
      "args": ["Target As Range"],
      "line": 1,
      "component_type": "document_module",
      "visibility": "Public",
      "has_parameters": true,
      "runnable": false,
      "reason_not_runnable": "has_parameters",
      "run_command": null
    }
  ],
  "suggestions": [
    {
      "title": "Run the default entrypoint",
      "command": "xlflow run Main.GenerateQRCode --session --json"
    }
  ]
}
```

### Field Reference

Each macro entry in the `macros` array includes:

| Field                 | Type       | Description                                                                              |
| --------------------- | ---------- | ---------------------------------------------------------------------------------------- |
| `module`              | `string`   | VBA component (module) name.                                                             |
| `name`                | `string`   | Procedure name.                                                                          |
| `qualified_name`      | `string`   | Fully qualified name (`Module.Procedure`). Use this with `xlflow run`.                   |
| `kind`                | `string`   | `"sub"` or `"function"`.                                                                 |
| `args`                | `string[]` | Raw parameter declarations (e.g. `"path As String"`).                                    |
| `line`                | `number`   | 1-based line number in the source module.                                                |
| `component_type`      | `string`   | One of: `standard_module`, `class_module`, `document_module`, `userform`, `unknown`.     |
| `visibility`          | `string`   | Declared visibility: `"Public"`, `"Private"`, `"Friend"`, or `"Implicit"`.               |
| `has_parameters`      | `boolean`  | Whether the procedure declares parameters.                                               |
| `runnable`            | `boolean`  | Whether the macro can be run directly via `xlflow run`.                                  |
| `reason_not_runnable` | `string?`  | If not runnable, why: `has_parameters`, `event_procedure`, `unsupported_component_type`. |
| `run_command`         | `string?`  | Ready-to-use `xlflow run` command (only present when `runnable` is `true`).              |

### Top-level Fields

| Field           | Type       | Description                                                            |
| --------------- | ---------- | ---------------------------------------------------------------------- |
| `default_entry` | `string?`  | The `project.entry` from `xlflow.toml` if it matches a runnable macro. |
| `suggestions`   | `object[]` | Suggested next commands, each with `title` and `command`.              |

### Runnable Determination

A macro is considered **runnable** when **all** of the following are true:

- `has_parameters` is `false` (no parameters)
- Not an event procedure (e.g. `Workbook_Open`, `Worksheet_Change`, `Auto_Open`)
- `component_type` is `standard_module`, `class_module`, or `document_module` (not `userform` or `unknown`)

## Related

- [run](./run)
- [inspect](./inspect)
