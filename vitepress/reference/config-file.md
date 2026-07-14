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
# Excel bridge mode. Valid values: "auto", "dotnet".
bridge = "auto"

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

# Automatic backup retention is disabled by default.
# [backup.retention]
# enabled = false
# max_count = 20
# max_age_days = 30
# min_keep = 5
# max_total_size_mb = 2048

# Static analysis rules.
[lint]
# Disable specific lint rules by diagnostic ID.
disabled_rules = []

# VB020 unused-local-variable warnings are enabled by default.
# Add "VB020" to disabled_rules if a project intentionally keeps scratch locals.
#
# Optional project-wide lint rules. Uncomment individual rules to enable them.
# detect_scope_shadowing = true          # VB018
# detect_unused_private_procedures = true # VB021
# detect_nested_with_ambiguity = true    # VB027

# Runtime-risk analysis rules.
[analyze]
# Disable specific analyzer rules by diagnostic ID.
disabled_rules = []
```

## Section reference

### `[project]`

| Key     | Type   | Required | Default    | Description                                                                                   |
| ------- | ------ | -------- | ---------- | --------------------------------------------------------------------------------------------- |
| `name`  | string | no       | `sample`   | Display name for the project. Falls back to the workbook base name if omitted.                |
| `entry` | string | **yes**  | `Main.Run` | Qualified name of the macro executed by `xlflow run` when no positional argument is supplied. |

### `[excel]`

| Key              | Type   | Required | Default           | Description                                                                                 |
| ---------------- | ------ | -------- | ----------------- | ------------------------------------------------------------------------------------------- |
| `path`           | string | **yes**  | `build/Book.xlsm` | Workbook, binary workbook, or add-in path. May be relative to the project root or absolute. |
| `visible`        | bool   | no       | `false`           | Whether the Excel application window is shown during automation.                            |
| `display_alerts` | bool   | no       | `false`           | Whether Excel shows its own alert dialogs (e.g. overwrite confirmation).                    |
| `bridge`         | string | no       | `"auto"`          | Excel bridge mode. Valid values are `"auto"` and `"dotnet"`.                                |

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

\* `xlflow new` defaults to `"sidecar"`. `xlflow init` defaults to `"frm"` so that existing code inside `.frm` files remains authoritative. Use `xlflow init --userform-code-source sidecar` for imported workbooks that should start with sidecar code and Designer specs.

### `[backup.retention]`

| Key                 | Type | Required | Default | Description                                                                 |
| ------------------- | ---- | -------- | ------- | --------------------------------------------------------------------------- |
| `enabled`           | bool | no       | `false` | Enable automatic pruning after successful backup-producing operations.      |
| `max_count`         | int  | no       | `20`    | Keep total valid backups within this count. `0` disables the count limit.   |
| `max_age_days`      | int  | no       | `30`    | Delete valid backups older than this many days. `0` disables the age limit. |
| `min_keep`          | int  | no       | `5`     | Always protect the newest valid backups from all automatic limits.          |
| `max_total_size_mb` | int  | no       | `2048`  | Keep total backup size within this decimal MB limit. `0` disables it.       |

Automatic retention is disabled by default and only affects backups for the configured `[excel].path`. Manual `xlflow backup prune` does not require `enabled = true`.

Negative numeric values are configuration errors. `min_keep > max_count` is also invalid when `max_count` is greater than zero. If all limits are disabled, automatic pruning performs no deletion. Invalid entries and legacy directories without metadata are skipped, not deleted.

### `[lint]`

| Key              | Type     | Required | Default | Description                                       |
| ---------------- | -------- | -------- | ------- | ------------------------------------------------- |
| `disabled_rules` | string[] | no       | `[]`    | Disable configurable lint rules by diagnostic ID. |

Legacy per-rule booleans such as `forbid_select = false` remain accepted for compatibility, but xlflow emits a deprecation warning. Prefer `disabled_rules = ["VB002"]`.

`VB020` unused-local-variable warnings are enabled by default and can be disabled with `disabled_rules = ["VB020"]`. Other project-wide lint rules such as `detect_unused_private_procedures = true` (`VB021`) remain disabled by default. New `xlflow.toml` files include commented examples so projects can opt in deliberately.

Configurable lint rule IDs:

| ID      | Legacy key                           |
| ------- | ------------------------------------ |
| `VB001` | `require_option_explicit`            |
| `VB002` | `forbid_select`                      |
| `VB003` | `forbid_activate`                    |
| `VB004` | `forbid_on_error_resume_next`        |
| `VB005` | `detect_implicit_variant`            |
| `VB006` | `forbid_public_module_fields`        |
| `VB007` | `forbid_interactive_input`           |
| `VB018` | `detect_scope_shadowing`             |
| `VB019` | `detect_multiple_declarator_clarity` |
| `VB020` | `detect_unused_local_variables`      |
| `VB021` | `detect_unused_private_procedures`   |
| `VB022` | `detect_confusing_call_syntax`       |
| `VB023` | `detect_for_each_control_type`       |
| `VB026` | `detect_dangerous_resume`            |
| `VB027` | `detect_nested_with_ambiguity`       |

Safety diagnostics `VB008` through `VB015`, `VB028`, `VB029`, `VB031`, and `VB032` are always enabled and cannot be disabled with `disabled_rules`.

For local exceptions, keep the rule enabled and suppress a specific source line with an apostrophe comment:

```vb
' xlflow:disable-next-line VB002
Range("A1").Select
Range("A2").Select ' xlflow:disable-line VB002
```

Preflight-blocking lint diagnostics `VB008` through `VB015`, `VB028`, `VB029`, `VB031`, and `VB032` cannot be suppressed inline.

### `[analyze]`

| Key              | Type     | Required | Default | Description                                           |
| ---------------- | -------- | -------- | ------- | ----------------------------------------------------- |
| `disabled_rules` | string[] | no       | `[]`    | Disable configurable analyzer rules by diagnostic ID. |

Legacy per-rule booleans such as `forbid_unqualified_excel_objects = false` remain accepted for compatibility, but xlflow emits a deprecation warning. Prefer `disabled_rules = ["VBA205"]`.

Configurable analyzer rule IDs:

| ID       | Legacy key                              |
| -------- | --------------------------------------- |
| `VBA201` | `detect_range_find_nothing_check`       |
| `VBA202` | `detect_object_use_before_set`          |
| `VBA203` | `detect_application_state_restore`      |
| `VBA204` | `detect_error_handler_fallthrough`      |
| `VBA205` | `forbid_unqualified_excel_objects`      |
| `VBA206` | `detect_byref_argument_mismatch`        |
| `VBA207` | `detect_dictionary_collection_guard`    |
| `VBA208` | `detect_redim_preserve_dimension`       |
| `VBA209` | `detect_object_array_comparison`        |
| `VBA210` | `detect_function_return_path`           |
| `VBA211` | `detect_excel_object_member_mismatch`   |
| `VBA212` | `detect_non_short_circuit_object_guard` |

Analyzer diagnostics `VBA101` through `VBA106` are always enabled and cannot be disabled with `disabled_rules`.

Analyzer diagnostics can use the same inline suppression syntax:

```vb
' xlflow:disable-next-line VBA205
Range("A1").Value = 1
```

Preflight-blocking analyzer errors such as `VBA104`, `VBA105`, `VBA106`, and `VBA211` cannot be suppressed inline. Unknown inline IDs, unsupported IDs, and unused suppressions are reported as command warnings.

## Defaults differ between `new` and `init`

- **`xlflow new`** writes `[userform].code_source = "sidecar"`.
- **`xlflow init`** writes `[userform].code_source = "frm"` by default.
- **`xlflow init --userform-code-source sidecar`** writes `[userform].code_source = "sidecar"` and generates imported UserForm specs under `src/forms/specs`.
- **`xlflow form migrate sidecar`** converts an existing imported project from `frm` to `sidecar` after creating sidecar code and Designer specs.
- **`xlflow new`** defaults to `build/Book.xlsm` when the workbook argument is omitted or has no extension. Use an explicit `.xlam` filename to create an Excel add-in project or `.xlsb` for an Excel Binary Workbook VBA project.
- **`xlflow init`** preserves the copied workbook filename and extension, including `.xlam` and `.xlsb`.

All other sections use the same defaults regardless of how the project was created.

## Stable contract

For the authoritative specification of configuration keys, validation rules, and exit-code contracts, see `docs/specs/cli-contract.md`.
