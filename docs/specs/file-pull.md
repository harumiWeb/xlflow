# Saved-Workbook File Pull

## Command

```text
xlflow pull --backend file [--formulas] [--json]
```

`excel` is the default backend. Backend selection is explicit; failure of one
backend never invokes the other. `--backend file --session` is invalid.

## Authority and host requirements

The file backend reads the configured saved `.xlsm` package. It never launches
or attaches to Excel, COM, VBIDE, or the bridge. A matching
`.xlflow/session.json` record produces warning
`file_pull_live_session_ignored`; unsaved live workbook changes are not read.
Unreadable session metadata produces `file_pull_session_state_unavailable`
without changing saved-file authority.

Under WSL, this explicit pure-Go backend remains local; the default Excel
backend continues to delegate to Windows xlflow.

Successful JSON includes:

```json
{
  "pull": {
    "backend": "file",
    "source": "saved_workbook",
    "code_page": 932,
    "modules": {
      "standard": 1,
      "class": 1,
      "document": 2,
      "form": 0
    }
  }
}
```

## Supported project surface

The initial backend accepts `.xlsm` only and requires exactly one
`xl/vbaProject.bin`. Standard modules become `.bas`, class modules become
importable `.cls`, and document modules become body-only `.bas` source under
their configured roots. Source is decoded using the VBA project's declared
code page and written as UTF-8 without BOM with a final newline.

Configured Rubberduck `@Folder` behavior and `[vba.line_numbers]` removal apply
equally to file and Excel pulls. Unsafe names, path escapes, duplicate target
paths, malformed projects, and unsafe line-number transformations fail before
publication.

UserForms are not partially supported. Detection of any form rejects the whole
operation with `pull_userform_unsupported` before source mutation. The forms
root is otherwise unmanaged and is never deleted or rewritten.
The configured forms root must not contain, or be contained by, the module,
class, or workbook root; overlap fails before reconciliation so form artifacts
cannot be mistaken for stale managed source.

## Publication contract

Planning and validation complete before publication. Publication reconciles
managed `.bas` and `.cls` files under the module, class, and workbook roots.
Unrelated files and the forms root are retained. If writing or stale-file
removal fails, xlflow restores every affected prior file and reports
`pull_source_publish_failed`.

## Stable failures

| Error code                      | Meaning                                                |
| ------------------------------- | ------------------------------------------------------ |
| `pull_args_invalid`             | Unknown backend or incompatible `--session`.           |
| `workbook_format_unsupported`   | File backend was requested for a non-`.xlsm` workbook. |
| `pull_file_not_found`           | The configured workbook does not exist.                |
| `pull_file_unreadable`          | The saved workbook cannot be read.                     |
| `pull_vba_project_missing`      | The package has no `xl/vbaProject.bin`.                |
| `pull_vba_project_malformed`    | The package or VBA project is malformed.               |
| `pull_protected_project`        | The saved VBA project is protected.                    |
| `pull_userform_unsupported`     | At least one UserForm is present.                      |
| `pull_source_path_unsafe`       | A component cannot be mapped inside its managed root.  |
| `vba_line_number_safety_failed` | Generated line numbers cannot be removed safely.       |
| `pull_source_publish_failed`    | Transactional source publication failed.               |

`--formulas` runs only after successful VBA publication and retains its
existing saved-workbook formula authority and partial-failure contract.
