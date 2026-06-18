# Config File

xlflow reads `xlflow.toml` from the project root. It is the single source of truth for workbook paths, source directories, VBE behaviour, and static analysis rules.

## Full annotated example

This is the file produced by `xlflow new`. The comments are preserved so that each option is self-documenting.

```toml
# Project identity and entry point.
[project]
# Project name used in output messages. Falls back to the workbook base name.
name = "sample"
# Default macro invoked by xlflow run when no positional macro is given.
entry = "Main.Run"

# Excel automation settings.
[excel]
# Path to the workbook, relative to the project root or absolute.
path = "build/Book.xlsm"
# Make the Excel application window visible during automation.
visible = false
# Suppress Excel alert dialogs (e.g. overwrite confirmations).
display_alerts = false

# Source tree directories.
[src]
# Directory for standard .bas modules.
modules = "src/modules"
# Directory for class .cls modules.
classes = "src/classes"
# Directory for UserForm .frm files.
forms = "src/forms"
# Directory for workbook document module text.
workbook = "src/workbook"

# VBE component folder support (Rubberduck-style).
[vba]
# Enable @Folder("A.B") annotations and nested source paths.
folders = true
# How xlflow handles @Folder annotations during push.
# Valid values: "update", "preserve", "ignore".
#   "update"    – rewrite from source directory layout.
#   "preserve"  – keep existing annotations as-is.
#   "ignore"    – disable folder annotation read/write.
folder_annotation = "update"
# Automatically assign default folder annotations based on source paths.
default_component_folders = true

# UserForm source mode.
[userform]
# Where UserForm code-behind lives in the source tree.
# Valid values: "frm", "sidecar".
#   "frm"     – code is kept inside the exported .frm file.
#   "sidecar" – code is split into src/forms/code/<FormName>.bas.
code_source = "sidecar"

# Static analysis rules.
[lint]
# Require Option Explicit in every module.
require_option_explicit = true
# Forbid Select / Activate patterns.
forbid_select = true
# Forbid Activate usage.
forbid_activate = true
# Forbid On Error Resume Next.
forbid_on_error_resume_next = true
# Detect implicitly typed Variant variables.
detect_implicit_variant = true
# Forbid public fields in standard modules.
forbid_public_module_fields = true
# Forbid interactive input (MsgBox, InputBox, etc.) in headless runs.
forbid_interactive_input = true

# Runtime-risk analysis rules.
[analyze]
# Detect Range.Find results used without a Nothing check.
detect_range_find_nothing_check = true
# Detect object variables used before an obvious Set assignment.
detect_object_use_before_set = true
# Detect Application state changes without an obvious restore path.
detect_application_state_restore = true
# Detect procedures that can fall through into an error handler.
detect_error_handler_fallthrough = true
# Forbid unqualified Range/Cells/Rows/Columns access.
forbid_unqualified_excel_objects = true
# Detect likely ByRef argument type mismatches.
detect_byref_argument_mismatch = false
# Detect Dictionary/Collection access without an obvious guard.
detect_dictionary_collection_guard = false
# Detect ReDim Preserve usage on multi-dimensional arrays.
detect_redim_preserve_dimension = true
# Detect object or array comparison mistakes.
detect_object_array_comparison = true
# Detect functions that may exit without assigning their return value.
detect_function_return_path = false
# Detect known Excel object/member mismatches.
detect_excel_object_member_mismatch = true
```

## Section reference

### `[project]`

| Key     | Type   | Required | Default    | Description                                                                                   |
| ------- | ------ | -------- | ---------- | --------------------------------------------------------------------------------------------- |
| `name`  | string | no       | `sample`   | Display name for the project. Falls back to the workbook base name if omitted.                |
| `entry` | string | **yes**  | `Main.Run` | Qualified name of the macro executed by `xlflow run` when no positional argument is supplied. |

### `[excel]`

| Key              | Type   | Required | Default           | Description                                                              |
| ---------------- | ------ | -------- | ----------------- | ------------------------------------------------------------------------ |
| `path`           | string | **yes**  | `build/Book.xlsm` | Workbook path. May be relative to the project root or absolute.          |
| `visible`        | bool   | no       | `false`           | Whether the Excel application window is shown during automation.         |
| `display_alerts` | bool   | no       | `false`           | Whether Excel shows its own alert dialogs (e.g. overwrite confirmation). |

### `[src]`

| Key        | Type   | Required | Default        | Description                                       |
| ---------- | ------ | -------- | -------------- | ------------------------------------------------- |
| `modules`  | string | no       | `src/modules`  | Root directory for standard `.bas` modules.       |
| `classes`  | string | no       | `src/classes`  | Root directory for class `.cls` modules.          |
| `forms`    | string | no       | `src/forms`    | Root directory for UserForm `.frm` files.         |
| `workbook` | string | no       | `src/workbook` | Root directory for workbook document module text. |

When `[vba].folders = true`, files may be nested under these roots according to `@Folder` annotations.

### `[vba]`

| Key                         | Type   | Required | Default    | Description                                                                                    |
| --------------------------- | ------ | -------- | ---------- | ---------------------------------------------------------------------------------------------- |
| `folders`                   | bool   | no       | `true`     | Enable Rubberduck-style `@Folder("A.B")` annotations and nested source paths.                  |
| `folder_annotation`         | string | no       | `"update"` | How `push` treats `@Folder` annotations.<br>Valid values: `"update"`, `"preserve"`, `"ignore"` |
| `default_component_folders` | bool   | no       | `true`     | Automatically assign default folder annotations based on the relative source path.             |

### `[userform]`

| Key           | Type   | Required | Default       | Description                                                                                |
| ------------- | ------ | -------- | ------------- | ------------------------------------------------------------------------------------------ |
| `code_source` | string | no       | `"sidecar"`\* | Where UserForm code-behind lives in the source tree.<br>Valid values: `"frm"`, `"sidecar"` |

\* `xlflow new` defaults to `"sidecar"`. `xlflow init` defaults to `"frm"` so that existing code inside `.frm` files remains authoritative.

### `[lint]`

| Key                           | Type | Required | Default | Description                                                             |
| ----------------------------- | ---- | -------- | ------- | ----------------------------------------------------------------------- |
| `require_option_explicit`     | bool | no       | `true`  | Require `Option Explicit` in every module.                              |
| `forbid_select`               | bool | no       | `true`  | Forbid `Select` / `Activate` patterns.                                  |
| `forbid_activate`             | bool | no       | `true`  | Forbid `Activate` usage.                                                |
| `forbid_on_error_resume_next` | bool | no       | `true`  | Forbid `On Error Resume Next`.                                          |
| `detect_implicit_variant`     | bool | no       | `true`  | Detect implicitly typed `Variant` variables.                            |
| `forbid_public_module_fields` | bool | no       | `true`  | Forbid public fields in standard modules.                               |
| `forbid_interactive_input`    | bool | no       | `true`  | Forbid interactive input (`MsgBox`, `InputBox`, etc.) in headless runs. |

### `[analyze]`

| Key                                   | Type | Required | Default | Description                                                        |
| ------------------------------------- | ---- | -------- | ------- | ------------------------------------------------------------------ |
| `detect_range_find_nothing_check`     | bool | no       | `true`  | Detect `Range.Find` results used without a `Nothing` check.        |
| `detect_object_use_before_set`        | bool | no       | `true`  | Detect object variables used before an obvious `Set` assignment.   |
| `detect_application_state_restore`    | bool | no       | `true`  | Detect Application state changes without an obvious restore path.  |
| `detect_error_handler_fallthrough`    | bool | no       | `true`  | Detect normal execution falling through into error-handler labels. |
| `forbid_unqualified_excel_objects`    | bool | no       | `true`  | Detect unqualified `Range`, `Cells`, `Rows`, and `Columns` access. |
| `detect_byref_argument_mismatch`      | bool | no       | `false` | Detect likely ByRef argument type mismatch candidates.             |
| `detect_dictionary_collection_guard`  | bool | no       | `false` | Detect Dictionary/Collection lookup without an obvious guard.      |
| `detect_redim_preserve_dimension`     | bool | no       | `true`  | Detect risky multi-dimensional `ReDim Preserve` usage.             |
| `detect_object_array_comparison`      | bool | no       | `true`  | Detect object or array comparison mistakes.                        |
| `detect_function_return_path`         | bool | no       | `false` | Detect functions that may exit without assigning a return value.   |
| `detect_excel_object_member_mismatch` | bool | no       | `true`  | Detect known Excel object/member mismatches.                       |

## Defaults differ between `new` and `init`

- **`xlflow new`** writes `[userform].code_source = "sidecar"`.
- **`xlflow init`** writes `[userform].code_source = "frm"`.

All other sections use the same defaults regardless of how the project was created.

## Stable contract

For the authoritative specification of configuration keys, validation rules, and exit-code contracts, see `docs/specs/cli-contract.md`.
