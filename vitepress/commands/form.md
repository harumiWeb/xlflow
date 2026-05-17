# xlflow form

Manage UserForms through Designer snapshots, rebuilds, and image export.

## Usage

```bash
xlflow form snapshot <name> --out <path>
xlflow form build <spec> [--overwrite]
xlflow form export-image <name> --out <png>
```

## Options and Arguments

| Option / argument      | Description                                                                   | Default |
| ---------------------- | ----------------------------------------------------------------------------- | ------- |
| `snapshot <name>`      | Save Designer state as JSON or YAML.                                          | -       |
| `build <spec>`         | Create or update a UserForm from a saved spec.                                | -       |
| `export-image <name>`  | Render a runtime UserForm to PNG.                                             | -       |
| `--out <path>`         | Output path for snapshots or images.                                          | -       |
| `--overwrite`          | Allow replacing an existing UserForm on build.                                | false   |
| `--session`            | Operate against the managed live session workbook.                            | false   |
| `--no-save`            | Leave session-backed build changes unsaved until `xlflow save`.               | false   |
| `--initializer <mode>` | Control initializer execution for image export.                               | default |
| `--keepalive`          | Write periodic progress heartbeat lines to stderr for long Excel-backed work. | false   |

## Examples

```bash
xlflow form snapshot CalendarForm --out src/forms/specs/CalendarForm.yaml --json
xlflow form build src/forms/specs/CalendarForm.yaml --overwrite --json
xlflow form export-image CalendarForm --out artifacts/CalendarForm.png --json
```

## Notes

::: important
The canonical Designer source is `src/forms/specs/*.yaml` or `*.json`; sidecar code lives separately under `src/forms/code/`.
:::

::: warning
`form export-image` depends on desktop Excel GUI behavior and may execute `UserForm_Initialize` depending on initializer settings.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "form build",
  "form": "CalendarForm",
  "designer": "src/forms/specs/CalendarForm.yaml",
  "overwritten": true
}
```

## Related

- [list](./list)
- [inspect form](./inspect)
- [UserForm guide](../guides/userform-development)
