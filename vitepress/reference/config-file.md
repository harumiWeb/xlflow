# Config File

## Formatter settings

The `[fmt]` table controls the source formatter. These options are read by `fmt` and the VS Code formatting action:

| Key                   | Meaning                                                              |
| --------------------- | -------------------------------------------------------------------- |
| `operator_spacing`    | Normalize spacing around operators.                                  |
| `declaration_spacing` | Keep declaration blocks separated consistently.                      |
| `keyword_casing`      | Normalize VBA keyword casing.                                        |
| `builtin_casing`      | Normalize curated built-in names without recasing user declarations. |

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

# Optional Erl instrumentation. When enabled, push adds temporary physical
# prepared-import-source line numbers to imported procedure statements; tracked source stays unnumbered.
# [vba.line_numbers]
# enabled = true

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

# Release build source filtering. This affects `build` only; `push` and `pack`
# always use the complete source tree.
# [build]
# Exclude project-root-relative doublestar glob patterns from `xlflow build`.
# exclude = ["src/modules/Tests/**"]

# Source-preflight diagnostic waivers.
[preflight]
# Diagnostics remain enabled and visible when allowed here; only their
# source-preflight blocking effect is waived. Excel/VBE compilation may still fail.
allowed_diagnostics = []

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

# Optional local procedure-name constant check (VB044).
# [lint.procedure_name_constant]
# enabled = true
# constant_name = "PROCEDURE_NAME"

# Runtime-risk analysis rules.
[analyze]
# Disable specific analyzer rules by diagnostic ID.
disabled_rules = []

# Procedure complexity metrics are independent from lint/analyze diagnostics.
[metrics]
exclude = []

[metrics.thresholds]
cyclomatic_complexity = 0
max_nesting_depth = 0
statement_count = 0
source_line_count = 0
branch_count = 0
loop_count = 0
goto_count = 0
exit_point_count = 0
parameter_count = 0
byref_parameter_count = 0
local_variable_count = 0
call_fan_out = 0

[metrics.hotspots]
procedure_top_n = 0
module_top_n = 0
procedure_score_threshold = 0
module_score_threshold = 0
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

### `[vba.line_numbers]`

| Key       | Type | Required | Default | Description                                                          |
| --------- | ---- | -------- | ------- | -------------------------------------------------------------------- |
| `enabled` | bool | no       | `false` | Opt in to temporary `Erl` line-number instrumentation during `push`. |

Set `enabled = true` only when runtime diagnostics need meaningful `Erl` values. `push` adds labels only to its temporary import copies, so the tracked source files are never rewritten. Folder annotations are updated before labels are generated, so labels use the physical line number of the prepared import source. They use fixed-width space padding and two spaces before the statement; they never use a colon (for example, ` 7  Debug.Print value`).

`pull` removes only labels that exactly match this generated format, keeping exported source unnumbered. To avoid changing intentional VBA control flow, xlflow stops the affected `push` or `pull` without transforming it when it finds pre-existing or mismatched numeric labels, or numeric `GoTo`, `GoSub`, or `Resume` targets. This is configuration-only behavior: `push` and `pull` have no line-number CLI flags.

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

### `[build]`

| Key       | Type     | Required | Default | Description                                                                                          |
| --------- | -------- | -------- | ------- | ---------------------------------------------------------------------------------------------------- |
| `exclude` | string[] | no       | `[]`    | Project-root-relative `doublestar` glob patterns excluded from `xlflow build` source selection only. |

Each pattern is normalized to `/`; Windows and WSL separators match identically. Absolute paths and patterns that traverse outside the project root are invalid. `xlflow build` reports unmatched patterns as `build_exclude_unmatched` warnings. A matching UserForm artifact excludes the whole form component, including its related `.frx`, sidecar code, and persisted spec files. This setting does not change source selection for `push` or `pack`.

### `[preflight]`

| Key                   | Type     | Required | Default | Description                                                                                  |
| --------------------- | -------- | -------- | ------- | -------------------------------------------------------------------------------------------- |
| `allowed_diagnostics` | string[] | no       | `[]`    | Let listed registry diagnostics pass source preflight while keeping the diagnostics enabled. |

Entries are trimmed, uppercased, deduplicated in first-occurrence order, and
must name rules whose diagnostic catalog metadata says
`preflight_blocking: true`. Unknown IDs, non-blocking IDs, and non-registry
integrity codes such as `FRM...` or `UFY...` are configuration errors. There is
no wildcard.

```toml
[preflight]
allowed_diagnostics = ["VB052"]
```

This setting changes only whether workbook automation proceeds. On each
supported `lint`, `analyze`, or LSP surface, the diagnostic remains an `error`
and cannot become inline-suppressible. `lint` and `analyze` retain their normal
CLI exit behavior; LSP publishes diagnostics and does not participate in the
CLI exit-status contract. When xlflow applies waivers, command output includes
aggregated `preflight_diagnostic_allowed` warnings, one per distinct
diagnostic ID with occurrence counts aggregated within each warning, because
Excel/VBE compilation may still fail. Keep the list empty in CI unless the
project has explicitly reviewed and accepted each waiver.

### `[lint]`

| Key              | Type     | Required | Default | Description                                       |
| ---------------- | -------- | -------- | ------- | ------------------------------------------------- |
| `disabled_rules` | string[] | no       | `[]`    | Disable configurable lint rules by diagnostic ID. |

Legacy per-rule booleans such as `forbid_select = false` remain accepted for compatibility, but xlflow emits a deprecation warning. Prefer `disabled_rules = ["VB002"]`.

`VB020` unused-local-variable warnings are enabled by default and can be disabled with `disabled_rules = ["VB020"]`. Other project-wide lint rules such as `detect_unused_private_procedures = true` (`VB021`) remain disabled by default. New `xlflow.toml` files include commented examples so projects can opt in deliberately.

`VB044` is disabled by default and uses the nested `[lint.procedure_name_constant]` table rather than a legacy boolean. When enabled, `constant_name` is required and must be a VBA identifier. xlflow compares matching procedure-local constants case-insensitively by name, but requires their direct string-literal values to match the enclosing procedure name exactly. `disabled_rules = ["VB044"]` takes precedence over `enabled = true`.

The generated [diagnostic catalog](./diagnostics) is the authoritative list of
configurable IDs, configuration keys, defaults, and non-configurable safety
diagnostics. The same metadata is available from `xlflow rules --json`.
`disabled_rules` remains case-insensitive, rejects unknown or non-configurable
IDs, removes duplicates, and takes precedence over an enabled legacy boolean or
the nested `VB044` setting.

For local exceptions, keep the rule enabled and suppress a specific source line with an apostrophe comment:

```vb
' xlflow:disable-next-line VB002
Range("A1").Select
Range("A2").Select ' xlflow:disable-line VB002
```

Rules marked `inline_suppressible: false` in the catalog cannot be suppressed
inline. This includes every preflight-blocking lint diagnostic.

### `[lint.procedure_name_constant]`

| Key             | Type   | Required | Default | Description                                                              |
| --------------- | ------ | -------- | ------- | ------------------------------------------------------------------------ |
| `enabled`       | bool   | no       | `false` | Enable `VB044` checks for existing matching procedure-local constants.   |
| `constant_name` | string | yes\*    | —       | Case-insensitive local constant name to check, such as `PROCEDURE_NAME`. |

\* Required when `enabled = true`; the value must be a VBA identifier. The rule supports `Sub`, `Function`, `Property Get`, `Property Let`, and `Property Set` in standard, class, document, and UserForm code. It reports only direct string-literal values and leaves missing or expression-based constants alone.

### `[analyze]`

| Key                                       | Type     | Required | Default | Description                                                                                                                      |
| ----------------------------------------- | -------- | -------- | ------- | -------------------------------------------------------------------------------------------------------------------------------- |
| `disabled_rules`                          | string[] | no       | `[]`    | Disable configurable analyzer rules by diagnostic ID.                                                                            |
| `detect_risky_module_state`               | bool     | no       | `false` | Opt in to `VBA240` project-wide module-state coupling analysis and read/write metrics.                                           |
| `detect_redim_preserve_in_loops`          | bool     | no       | `true`  | Compatibility switch for default-enabled `VBA241` repeated `ReDim Preserve` analysis inside loops.                               |
| `detect_expensive_full_range_operations`  | bool     | no       | `false` | Opt in to `VBA242` full-row, full-column, full-sheet, and unbounded `UsedRange` operation analysis.                              |
| `detect_value2_performance_opportunities` | bool     | no       | `false` | Opt in to `VBA243` suggestions to use `Range.Value2` for bulk or repeated transfers when Date/Currency coercion is not required. |
| `detect_procedure_call_cycles`            | bool     | no       | `true`  | Compatibility switch for default-enabled `VBA244` project-wide recursive and cyclic procedure dependency analysis.               |
| `detect_unsafe_http_configuration`        | bool     | no       | `true`  | Compatibility switch for default-enabled `VBA246` HTTP transport-security analysis.                                              |
| `detect_missing_http_timeout`             | bool     | no       | `true`  | Compatibility switch for default-enabled `VBA247` HTTP timeout reliability analysis.                                             |
| `development_http_origins`                | string[] | no       | `[]`    | Exact plain-HTTP origins exempted only from `VBA246` plain-HTTP credential findings.                                             |

Legacy per-rule booleans such as `forbid_unqualified_excel_objects = false` remain accepted for compatibility, but xlflow emits a deprecation warning. Prefer `disabled_rules = ["VBA205"]`. `detect_risky_module_state = true` is the opt-in compatibility key for `VBA240`; disable it with `disabled_rules = ["VBA240"]`. `detect_redim_preserve_in_loops = false` is the compatibility key for disabling `VBA241`; prefer `disabled_rules = ["VBA241"]`. `detect_expensive_full_range_operations = true` is the opt-in compatibility key for `VBA242`; prefer enabling it explicitly only for Excel projects and use `disabled_rules = ["VBA242"]` to suppress it under a project policy. `detect_value2_performance_opportunities = true` is the opt-in compatibility key for `VBA243`; prefer enabling it explicitly only for Excel projects and use `disabled_rules = ["VBA243"]` to suppress it under a project policy. `detect_procedure_call_cycles = false` is the compatibility key for disabling `VBA244`; prefer `disabled_rules = ["VBA244"]` so the policy is explicit. `detect_unsafe_http_configuration` and `detect_missing_http_timeout` are compatibility switches for `VBA246` and `VBA247`; prefer `disabled_rules` for policy suppression.

`development_http_origins` accepts only exact absolute origins of the form `http://host[:port]`. It rejects credentials, paths (including a trailing slash), queries, fragments, wildcards, and HTTPS values. Host names are case-normalized, IPv6 is canonicalized, and the default port `:80` is removed. The exemption applies only to `plain_http_credentials`; loopback origins are exempt without configuration.

The generated [diagnostic catalog](./diagnostics) is the authoritative list of
analyzer configuration keys, defaults, and always-enabled diagnostics.

Analyzer diagnostics can use the same inline suppression syntax:

```vb
' xlflow:disable-next-line VBA205
Range("A1").Value = 1
```

Rules marked `inline_suppressible: false` in the catalog cannot be suppressed
inline. This includes every preflight-blocking analyzer error. Unknown inline
IDs, unsupported IDs, and unused suppressions are reported as command warnings.

### `[metrics]`

| Key       | Type     | Required | Default | Description                                                                                     |
| --------- | -------- | -------- | ------- | ----------------------------------------------------------------------------------------------- |
| `exclude` | string[] | no       | `[]`    | Project-root-relative `doublestar` globs excluded from `xlflow metrics` after source discovery. |

`exclude` is normalized to `/`, deduplicated, and evaluated in stable order.
Absolute patterns and patterns that traverse outside the project root are
invalid. A valid pattern that matches no source produces a
`metrics_exclude_unmatched` warning. The setting affects only `metrics`; it is
not shared with `[build].exclude`, `[lint].disabled_rules`, or
`[analyze].disabled_rules`.

### `[metrics.thresholds]`

| Key                     | Type | Required | Default | Description                                                             |
| ----------------------- | ---- | -------- | ------- | ----------------------------------------------------------------------- |
| `cyclomatic_complexity` | int  | no       | `0`     | Report `MX001` when the procedure value is greater than this threshold. |
| `max_nesting_depth`     | int  | no       | `0`     | Report `MX001` when the procedure value is greater than this threshold. |
| `statement_count`       | int  | no       | `0`     | Report `MX001` when the procedure value is greater than this threshold. |
| `source_line_count`     | int  | no       | `0`     | Report `MX001` when the procedure value is greater than this threshold. |
| `branch_count`          | int  | no       | `0`     | Report `MX001` when the procedure value is greater than this threshold. |
| `loop_count`            | int  | no       | `0`     | Report `MX001` when the procedure value is greater than this threshold. |
| `goto_count`            | int  | no       | `0`     | Report `MX001` when the procedure value is greater than this threshold. |
| `exit_point_count`      | int  | no       | `0`     | Report `MX001` when the procedure value is greater than this threshold. |
| `parameter_count`       | int  | no       | `0`     | Report `MX001` when the procedure value is greater than this threshold. |
| `byref_parameter_count` | int  | no       | `0`     | Report `MX001` when the procedure value is greater than this threshold. |
| `local_variable_count`  | int  | no       | `0`     | Report `MX001` when the procedure value is greater than this threshold. |
| `call_fan_out`          | int  | no       | `0`     | Report `MX001` when the procedure value is greater than this threshold. |

The twelve names are the v1 metric schema. Omitted and zero values disable a
threshold; positive values use a strict `value > threshold` comparison, so an
equal value passes. Negative values, non-integers, unknown keys, and malformed
tables are configuration errors. Thresholds are evaluated after the complete
metrics collection and do not change raw values. Each exceeded
procedure/metric pair produces one metrics-specific warning diagnostic with
code `MX001`; it is not a static-analysis rule, cannot be inline-suppressed,
and is not controlled by `[analyze].disabled_rules`. Without enabled
thresholds, `xlflow metrics` exits `0`; one or more `MX001` diagnostics retain
the full metrics result and exit `1` with `error.code =
"metrics_threshold_exceeded"`.

### `[metrics.hotspots]`

| Key                         | Type  | Required | Default | Description                                                                                |
| --------------------------- | ----- | -------- | ------- | ------------------------------------------------------------------------------------------ |
| `procedure_top_n`           | int   | no       | `0`     | Select the top N ranked procedure hotspots; `0` disables top-N selection.                  |
| `module_top_n`              | int   | no       | `0`     | Select the top N ranked module hotspots; `0` disables top-N selection.                     |
| `procedure_score_threshold` | float | no       | `0`     | Select procedure scores greater than or equal to this percentage (`1..100`); `0` disables. |
| `module_score_threshold`    | float | no       | `0`     | Select module scores greater than or equal to this percentage (`1..100`); `0` disables.    |

Hotspots are emitted under `metrics.hotspots` by `xlflow metrics` even when no
selector is enabled. Scores use the versioned
`percentile_equal_weight_v1` model and retain `raw_signals` plus
`normalized_signals`; procedure and module cohorts are ranked independently.
Top-N and score-threshold selections are unioned. Top-N-only `MX002` entries
are informational; threshold-selected entries are warning-level and return
`error.code = "metrics_hotspot_threshold_exceeded"` with exit `1`, while
retaining the complete metrics payload. Negative top-N values, non-finite
numbers, unknown keys, malformed tables, and percentages outside `0..100` are
configuration errors (exit `2`). The signal vocabulary and deterministic
ranking contract are defined in
`docs/specs/vba-procedure-and-module-hotspots.md`.

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
