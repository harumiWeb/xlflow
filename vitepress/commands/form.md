# xlflow form

Manage UserForms through sidecar scaffolds, Designer snapshots, rebuilds, and image export.

## Usage

```bash
xlflow form new <name>
xlflow form migrate sidecar [FormName] [--overwrite]
xlflow form snapshot <name> --out <path>
xlflow form build <spec> [--overwrite]
xlflow form export-image <name> --out <png>
```

## Options and Arguments

| Option / argument      | Description                                                     | Default |
| ---------------------- | --------------------------------------------------------------- | ------- |
| `new <name>`           | Create sidecar UserForm code/spec source files.                 | -       |
| `migrate sidecar`      | Convert frm-mode UserForms to sidecar code and Designer specs.  | -       |
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

`form migrate sidecar` converts existing tracked `src/forms/*.frm` UserForms to the sidecar layout. It extracts code-behind to `src/forms/code/<Name>.bas`, writes a non-executing Designer spec to `src/forms/specs/<Name>.yaml`, and switches `[userform].code_source` to `"sidecar"` after all selected forms succeed. Pass a form name to migrate one form. Existing Designer spec files always require `--overwrite`; existing sidecar code can be skipped only when it already matches the extracted `.frm` code.

Because the Designer spec is captured from the configured workbook, migration refuses to run when source files are newer than the workbook. Run `xlflow push` to apply source-only UserForm edits first, or `xlflow pull` if the workbook should remain authoritative.

## Examples

```bash
xlflow form new CustomerForm --json
xlflow form migrate sidecar --json
xlflow form migrate sidecar CalendarForm --overwrite --json
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
> `form migrate sidecar` keeps `.frm` and `.frx` artifacts in place. After migration, edit Designer structure in `src/forms/specs/*.yaml` and code-behind in `src/forms/code/*.bas`; treat `.frm` / `.frx` as generated artifacts.
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
  "command": "form migrate sidecar",
  "source": {
    "operation": "userform.migrate_sidecar",
    "code_source_before": "frm",
    "code_source_after": "sidecar",
    "forms": [
      {
        "name": "CalendarForm",
        "frm_path": "src/forms/CalendarForm.frm",
        "frx_path": "src/forms/CalendarForm.frx",
        "code_path": "src/forms/code/CalendarForm.bas",
        "spec_path": "src/forms/specs/CalendarForm.yaml"
      }
    ],
    "created": ["src/forms/code/CalendarForm.bas", "src/forms/specs/CalendarForm.yaml"],
    "updated": ["xlflow.toml"],
    "skipped": [],
    "config_path": "xlflow.toml",
    "requires_push": false
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
