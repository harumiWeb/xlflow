# xlflow form

Manage UserForms through sidecar scaffolds, Designer snapshots, rebuilds, and image export.

## Usage

```bash
xlflow form new <name>
xlflow form snapshot <name> --out <path>
xlflow form build <spec> [--overwrite]
xlflow form export-image <name> --out <png>
```

## Options and Arguments

| Option / argument      | Description                                                     | Default |
| ---------------------- | --------------------------------------------------------------- | ------- |
| `new <name>`           | Create sidecar UserForm code/spec source files.                 | -       |
| `snapshot <name>`      | Save Designer state as JSON or YAML.                            | -       |
| `build <spec>`         | Create or update a UserForm from a saved spec.                  | -       |
| `export-image <name>`  | Render a runtime UserForm to PNG.                               | -       |
| `--out <path>`         | Output path for snapshots or images.                            | -       |
| `--overwrite`          | Allow replacing an existing UserForm on build.                  | false   |
| `--session`            | Operate against the managed live session workbook.              | false   |
| `--no-save`            | Leave session-backed build changes unsaved until `xlflow save`. | false   |
| `--initializer <mode>` | Control initializer execution for image export.                 | default |

`form snapshot` reads the design-time Designer state without loading the form at runtime or running workbook VBA. Controls created only by runtime code are visible through runtime inspection or image export, not through snapshot.

`form new` is source-only and requires `[userform].code_source = "sidecar"`. It creates `[src].forms/code/<Name>.bas` and `[src].forms/specs/<Name>.yaml` (defaulting to `src/forms/...`); it does not create `.frm` or `.frx` artifacts.

## Examples

```bash
xlflow form new CustomerForm --json
xlflow form snapshot CalendarForm --out src/forms/specs/CalendarForm.yaml --json
xlflow form build src/forms/specs/CalendarForm.yaml --overwrite --json
xlflow form export-image CalendarForm --out artifacts/CalendarForm.png --json
```

## Notes

> [!IMPORTANT]
> The canonical Designer source is `src/forms/specs/*.yaml` or `*.json`; sidecar code lives separately under `src/forms/code/`.

> [!WARNING]
> `form new` refuses to overwrite an existing sidecar code or spec file. In `frm` code-source mode, use `form snapshot` / `form build` workflows or switch the project to sidecar mode intentionally.
>
> [!WARNING]
> `form export-image` depends on desktop Excel GUI behavior and may execute `UserForm_Initialize` depending on initializer settings.

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "form new",
  "source": {
    "created": ["src/forms/code/CustomerForm.bas", "src/forms/specs/CustomerForm.yaml"],
    "kind": "form",
    "name": "CustomerForm",
    "code_path": "src/forms/code/CustomerForm.bas",
    "spec_path": "src/forms/specs/CustomerForm.yaml",
    "code_source": "sidecar"
  }
}
```

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
