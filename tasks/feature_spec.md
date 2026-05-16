# xlflow Folder Structure Spec

## UserForm Phase 1 Warning Spec

## Phase 2 Goal

Make UserForm risk visible in existing workflows before dedicated form commands exist.

## Contract

- `pull`, `push`, `save`, and `session start|status|stop` add top-level `warnings` and `hints` when UserForms are detected.
- `inspect` remains file-based and detects UserForms from configured `src.forms` `.frm` files without opening Excel COM.
- Phase 1 does not add new `form` subcommands.

## Phase 2 Behavior

- Generic workbook/source detection warning uses `userform_state_partial` and explains that `.frm` text alone may not reflect layout, `.frx`, or Designer-backed state.
- Generic guidance hint uses `userform_planned_commands` and lists the spec-first workflow: `inspect form --designer`, `form snapshot --out src/forms/specs/<name>.yaml`, edit spec, `form build --overwrite`, and `form export-image`.
- `push --session --no-save` with detected UserForms adds `userform_unsaved_session_state` warning in addition to existing save-required state.
- `inspect` with detected source UserForms adds `userform_inspect_saved_file` warning so callers do not confuse saved workbook snapshots with live Designer/runtime form state.
- UserForm detection is best-effort for session lifecycle commands; inability to inspect forms must not fail `session`.

## UserForm Phase 2 List Forms Spec

## Goal

List workbook UserForms through Excel COM so agents can discover form names and expected source companion paths before deeper inspection commands exist.

## CLI Contract

- `xlflow list forms [--session] [--keepalive] [--keepalive-interval <duration>]`

## Behavior

- `list forms` opens the configured workbook through the normal session-aware Excel bridge.
- It inspects `VBProject.VBComponents` and returns only components with type `3` (`vbext_ct_MSForm`).
- It does not load forms at runtime and does not execute `UserForm_Initialize`.
- Each form result includes `name`, `component_type = "MSForm"`, `has_frx`, `source_path`, and optional `frx_path`.
- `source_path` and `frx_path` are expected project-relative source tree paths derived with the same folder-aware path resolution as `pull`.
- `has_frx` reports whether the expected `.frx` sibling currently exists in the source tree.
- VBProject access denial fails with `vbproject_access_denied` and a Trust Center hint.
- When a matching live session workbook is reused and newer than disk, the result includes the normal save-required metadata and warning.

## Verification

- CLI exposes `list forms` with `--session` and keepalive flags.
- Bridge args include workbook path, source roots, folder config, session metadata path, and project root for relative path resolution.
- PowerShell bridge returns filtered UserForm components plus expected `.frm` / `.frx` source paths.

## Goal

Support Rubberduck-compatible `@Folder(...)` annotations and nested source directories while preserving the existing type-specific `[src]` roots.

## Config Contract

```toml
[vba]
folders = true
folder_annotation = "update"
default_component_folders = true
```

- `folders`: enables folder-aware pull/push behavior.
- `folder_annotation`: one of `update`, `preserve`, `ignore`.
- `default_component_folders`: reserved compatibility flag; current typed `[src]` defaults remain the effective fallback roots when no annotation/path nesting exists.

## Behavior

- `pull` exports modules recursively under the configured root for each component type.
- When a valid `@Folder("Domain.Services")` annotation exists near the top of a component, `pull` writes the file under `.../Domain/Services/<Component>.<ext>`.
- Invalid or malformed annotations are ignored for path resolution.
- `pull` clears stale exported `.bas`, `.cls`, `.frm`, and `.frx` files under the configured source roots before re-export.
- `push` treats filesystem location as authoritative and rewrites folder annotations in temporary import copies when `folder_annotation = "update"`.
- `push` preserves annotations as-is when `folder_annotation = "preserve"` and does not read/write them when `folder_annotation = "ignore"`.
- Duplicate VBA component basenames anywhere in the recursive source tree fail `push` with `duplicate_module_name`.
- UserForm `.frm` and `.frx` companions must remain siblings after nested moves.

## UserForm Phase 3 Inspect Form Spec

### Goal

Inspect existing workbook UserForms through Excel COM and return structured control state for AI-agent review.

### Commands

- `xlflow inspect form <name> [--runtime|--designer|--both] [--initializer <method>] [--session] [--keepalive] [--keepalive-interval <duration>]`
- `xlflow inspect form UserForm1 --runtime --initializer InitializeForm --json`

### Behavior

- Default basis is `runtime`.
- `runtime` loads the form, inspects the loaded control tree, and unloads it before returning.
- `designer` inspects `VBProject.VBComponents(name).Designer` from the source workbook without loading the form at runtime.
- `runtime` and `both` execute against a temporary workbook copy created from the current source workbook state so form initialization does not mutate the source workbook or live session.
- `both` returns both snapshots in one response.
- `--initializer` is optional and is allowed only with `runtime` or `both`.
- When provided, xlflow invokes the named public form method with `ThisWorkbook` after runtime load and before control enumeration.
- Runtime inspection adds warnings that `UserForm_Initialize` ran and, when applicable, that the explicit initializer ran.
- The result includes top-level `target`, `session`, and `warnings` metadata like other workbook-backed commands.

### Output

- `inspect.target = "form"`
- `inspect.source = "excel_com"`
- Single-basis output uses `inspect.form`.
- `--both` uses `inspect.forms.runtime` and `inspect.forms.designer`.
- Form snapshots include `name`, `basis`, `caption`, optional `width` / `height`, `coordinate_system`, and `controls`.
- Designer inspect returns only true top-level controls at the root `controls` array; nested controls remain only under their parent container.
- Control snapshots include `name`, `type`, optional `prog_id`, `caption`, `text`, `value`, `left`, `top`, `width`, `height`, `tab_index`, `enabled`, `visible`, optional `selected_index`, optional `list`, and recursive `controls` for supported container controls.

### Errors

- `inspect_form_args_invalid`
- `form_not_found`
- `vbproject_access_denied`
- `designer_access_failed`
- `runtime_form_load_failed`
- `form_initializer_failed`
- `control_enumeration_failed`

## UserForm Phase 4 Snapshot Spec

### Goal

Persist a strict design-time UserForm snapshot as a stable JSON/YAML spec file for human review and later declarative workflows, without changing `inspect form --designer` away from its direct read-only VBIDE path.

### Commands

- `xlflow form snapshot <name> --out <path> [--session] [--keepalive] [--keepalive-interval <duration>]`
- `xlflow form snapshot UserForm1 --out src/forms/specs/UserForm1.yaml --session --json`

### Behavior

- Phase 4 is fixed to strict designer snapshot mode. There is no `--designer` flag, and `--runtime`, `--both`, and `--initializer` are not supported.
- `inspect form --designer` remains the direct read-only VBIDE inspection path.
- `form snapshot` opens a temporary workbook copy and runs an injected VBA helper to recover concrete control types for the persisted artifact.
- xlflow converts the returned designer snapshot into a persisted form spec in Go.
- `--out` is required.
- Output format is selected only by the `--out` extension.
- `.json` writes JSON.
- `.yaml` and `.yml` write YAML using the same spec fields.
- Any other extension fails with `form_snapshot_args_invalid` before Excel is opened.
- Snapshot can fail when the workbook's VBA project cannot execute the injected helper.
- `--session`, `--keepalive`, and `--keepalive-interval` behave the same as `inspect form`.
- Successful runs write the spec file, return normal workbook/session metadata, and expose the written path via top-level `output` metadata.
- Persisted `warnings` are limited to form-local snapshot warnings. Top-level operational warnings such as `save_required` stay in the command result instead of the saved artifact.
- Canonical source-controlled UserForm artifacts should live under `src/forms/specs/*.yaml`; exported `.frm` / `.frx` files are build or pull artifacts rather than the primary source of truth for Designer behavior.
- Phase 4 does not add new `.frx` parsing or unsupported-property detection beyond what the strict helper snapshot exposes.

### Output Spec

- Top-level persisted keys are `schemaVersion`, `kind`, `basis`, `coordinateSystem`, `form`, `controls`, and `warnings`.
- `schemaVersion = 1`
- `kind = "xlflow.userform"`
- `basis = "designer"`
- `coordinateSystem` comes from inspect `coordinate_system`.
- `form` contains `name`, optional legacy `caption` / `width` / `height`, optional raw `observed`, and optional rebuild-target `build`.
- `controls` is the canonical flat capture array. Each control has `id`, optional `parentId`, optional `zIndex`, the normalized camelCase control fields, and optional `observed`. Legacy nested `controls` input may still be accepted and normalized into the flat array.
- Explicit duplicate control `id` values are validation errors. They are not auto-corrected because `id` and `parentId` define the canonical tree.
- Control fields convert inspect snake_case keys to camelCase where needed: `prog_id -> progId`, `tab_index -> tabIndex`, `selected_index -> selectedIndex`.

### Errors

- `form_snapshot_args_invalid`
- `form_snapshot_write_failed`
- `form_not_found`
- `vbproject_access_denied`
- `designer_access_failed`
- `control_enumeration_failed`

### Snapshot validation target

- Workspace: `tmp_workspaces/user-form`
- Workbook: `build/Book.xlsm`
- Form: `UserForm1`
- JSON validation path: `xlflow form snapshot UserForm1 --out artifacts/UserForm1.json --json`
- YAML validation path: `xlflow form snapshot UserForm1 --out artifacts/UserForm1.yaml --json`

### Runtime inspection validation target

- Workspace: `tmp_workspaces/user-form`
- Workbook: `build/Book.xlsm`
- Form: `UserForm1`
- Runtime validation path: `xlflow inspect form UserForm1 --runtime --initializer InitializeForm --json`

## UserForm Phase 5 Form Export Image Spec

### Goal

Export a runtime-rendered workbook UserForm as a PNG image for visual verification without mutating the source workbook or attached live session.

### Commands

- `xlflow form export-image <name> --out <path.png> [--initializer <method>] [--overwrite] [--session] [--keepalive] [--keepalive-interval <duration>]`
- `xlflow form export-image UserForm1 --out artifacts/UserForm1.png --initializer InitializeForm --json`

### Behavior

- Phase 5 uses the `form` namespace: the CLI is `xlflow form export-image`.
- Capture basis is fixed to runtime.
- Runtime capture executes against a temporary workbook copy created from the current source workbook state, matching the safety model of `inspect form --runtime`.
- The command loads the target form in the temporary workbook, optionally invokes the named public initializer with `ThisWorkbook`, shows the form modeless with a unique caption token, captures the matching top-level window, then unloads the form and closes the temporary workbook copy.
- `--initializer` is optional and mirrors `inspect form --runtime --initializer`.
- `--out` is required.
- Output format is selected only by the `--out` extension and is PNG-only in Phase 5.
- Any non-`.png` output path fails with `unsupported_image_format` before Excel is opened.
- Existing output files fail with `output_file_exists` unless `--overwrite` is set.
- Parent directories for `--out` are created automatically when missing.
- Success returns normal workbook/session metadata plus top-level `target`, `forms`, `output`, and `warnings`.
- The command always warns that runtime capture executes `UserForm_Initialize`.
- The command always warns that runtime capture uses a temporary workbook copy.
- The command always warns that the feature is experimental and currently depends on Windows desktop Excel GUI behavior.
- Cleanup failures after a successful export are warnings, not fatal errors.

### Output

- `target.kind = file | live_session`
- `target.path = <workbook path>`
- `target.form = <form name>`
- `target.capture_state = "temporary_copy"`
- `forms.name = <form name>`
- `forms.basis = "runtime"`
- `forms.initializer = <method>` when provided
- `output.path = <resolved png path>`
- `output.format = "png"`
- `output.width_px` / `output.height_px` when image dimensions are available

### Errors

- `form_export_image_args_invalid`
- `unsupported_image_format`
- `output_file_exists`
- `form_not_found`
- `vbproject_access_denied`
- `runtime_form_load_failed`
- `form_initializer_failed`
- `window_not_found`
- `image_capture_failed`
- `temporary_component_cleanup_failed` as a warning when the main export succeeds

### Sample validation target

- Workspace: `tmp_workspaces/user-form`
- Workbook: `build/Book.xlsm`
- Form: `UserForm1`
- Validation path: `xlflow form export-image UserForm1 --out artifacts/UserForm1.png --initializer InitializeForm --json`

## UserForm Phase 6 Forms Package Extraction Spec

### Goal

Move UserForm-specific spec parsing, snapshot serialization, and internal types behind `internal/excel/forms` so the rest of xlflow stops depending on ad hoc raw payload handling.

### Behavior

- `internal/excel/forms` owns `FormSpec`, persisted snapshot read/write, inspect-snapshot conversion, and spec validation.
- Existing external behavior for `inspect form`, `form snapshot`, and `form export-image` remains unchanged.
- `internal/excel/form_snapshot.go` may remain as a compatibility wrapper, but new code should import `internal/excel/forms` directly.
- Form specs remain `schemaVersion = 1`, `kind = "xlflow.userform"`, `basis = "designer"`.

## UserForm Phase 7 Build / Apply Spec

### Goal

Create or rewrite Designer-backed workbook UserForms from persisted `xlflow.userform` specs without directly editing `.frx`.

### Commands

- `xlflow form build <spec.json|spec.yaml|spec.yml> [--overwrite] [--session] [--no-save] [--keepalive] [--keepalive-interval <duration>]`

### Behavior

- `form build` is the public command surface.
- `form apply` remains implemented but hidden because in-place Designer mutation is less stable than rebuild, is not maintained for sidecar-aware code-behind flows, and is a likely deletion candidate.
- Both code paths load and validate the spec in Go before Excel opens.
- Both code paths use the VBIDE Designer API from the PowerShell bridge; they do not parse or write `.frx` directly.
- `form build` creates a new UserForm from `form.name`.
- `form build` fails with `form_already_exists` when that form already exists, unless `--overwrite` is set.
- `form build --overwrite` exports the existing UserForm to a temporary backup, removes the existing UserForm component, saves the workbook, and recreates it from the spec.
- `form build --overwrite` must fail instead of deleting a same-name non-UserForm component.
- If overwrite rebuild fails after the deletion checkpoint, xlflow restores the original UserForm from the temporary backup and saves that restoration before returning failure.
- Public replacement workflow is `form build --overwrite`, which removes the existing component and recreates it from spec. The intended canonical Designer artifact is `src/forms/specs/*.yaml`; code-behind authority depends on `[userform].code_source`, with `src/forms/code/*.bas` canonical only in `sidecar` mode and embedded `.frm` code canonical in `frm` mode.
- In `sidecar` mode, tracked `.frm` embedded code is treated as a generated artifact and synchronized from `src/forms/code/*.bas` before `push` or `form build` opens Excel.
- v1 supported controls are `Label`, `TextBox`, `ComboBox`, `ListBox`, `CommandButton`, `CheckBox`, `OptionButton`, and `Frame`.
- Unsupported control types fail with `unsupported_form_control`.
- Unsupported property assignments are warnings, not fatal errors.
- Successful `form build` results should warn when the spec depends on weak Designer-backed fields: form-level `width` / `height` are best-effort, and design-time `ComboBox` / `ListBox` `list` / `selectedIndex` are observed-only for round-trip expectations even though xlflow still attempts to apply them.
- Parse failures should return `spec_parse_failed`; top-level schema failures should return `spec_schema_invalid`; structural validation failures should return `spec_validation_failed`.
- `form build` failure JSON should include top-level `spec.path`, `spec.format`, optional `spec.line`, optional `spec.column`, optional `spec.field`, and optional `spec.suggestion`.
- Both commands save by default.
- `--no-save` is allowed only with `--session`.
- `--overwrite --no-save` is invalid because Excel requires an intermediate save after removing the old UserForm and before recreating it.
- Successful `--session --no-save` runs must return normal save-required workbook/session metadata and warnings.

### Function / Type Contracts

- `buildForm(specPath string, opts FormWriteOptions) (FormWriteResult, error)`
- `applyForm(specPath string, opts FormWriteOptions) (FormWriteResult, error)`
- `loadFormSpec(specPath string) (FormSpec, SpecInput, error)`
- `validateFormWritePreflight(command string, cfg Config, formName string) error`
- `writeWorkbookForm(action string, spec FormSpec, opts FormWriteOptions) (FormWriteResult, error)`
- `exportUserFormBackup(name string) (UserFormBackup, error)`
- `restoreUserFormBackup(backup UserFormBackup) error`
- `createUserFormFromSpec(spec FormSpec) (CreatedUserForm, error)`
- `applyUserFormSpec(spec FormSpec) (UpdatedUserForm, error)`

### Type Notes

- `FormWriteOptions` contains `overwrite`, `session`, `noSave`, `keepalive`, and `keepaliveInterval`.
- `FormWriteResult` contains top-level `target`, `session`, optional `warnings` / `hints`, and `forms`.
- `forms` contains `name`, `basis`, `action`, `control_count`, `spec_path`, optional `caption`, optional `coordinate_system`, and `overwrite`.
- `UserFormBackup` identifies the temporary exported artifact path and original form name used for overwrite restore.
- `CreatedUserForm` / `UpdatedUserForm` identify the target form name plus any exported artifact sync metadata needed for result rendering.

### Error Contract

- `buildForm` returns `form_build_args_invalid` for option contract violations such as `--no-save` without `--session` and `--overwrite --no-save`.
- `applyForm` returns `form_apply_args_invalid` for apply-specific argument validation failures.
- Spec decode failures return `spec_parse_failed`, top-level schema failures return `spec_schema_invalid`, and structural validation failures return `spec_validation_failed`.
- Workbook write flows can return `form_already_exists`, `form_not_found`, `unsupported_form_control`, and `designer_write_failed`.

### Output

- Success uses top-level `target`, `session`, optional `warnings` / `hints`, and `forms`.
- `forms` includes `name`, `basis`, `action`, `control_count`, `spec_path`, optional `caption`, optional `coordinate_system`, and `overwrite` for build.

### Errors

- `form_build_args_invalid`
- `form_apply_args_invalid`
- `form_already_exists`
- `form_not_found`
- `unsupported_form_control`
- `designer_write_failed`

## Verification

- Config defaults and validation for `[vba]`.
- Recursive source fingerprint coverage.
- Annotation-aware component path resolution.
- Temporary import annotation rewrite from path.
- Recursive duplicate detection before Excel import.

# xlflow Performance Mode Spec

## Goal

Speed up repeated agent development loops while preserving the current safe defaults for CI and release workflows.

## CLI Contract

- `xlflow push --backup=always|never [--changed-only] [--fast] [--session] [--no-save]`
- `xlflow run [macro] [--direct] [--fast] [--session]`
- `xlflow runner install|remove|status`
- `xlflow session start|status|stop`
- `xlflow save --session`

## Behavior

- Default `push` still backs up, fully replaces non-document components, updates document modules, saves, and closes Excel.
- `push --backup=never` skips backup export.
- `push --changed-only` uses `.xlflow/state/push.json`; unchanged source skips Excel/VBIDE import, changed or missing state falls back to full push.
- `push --fast` expands to `--backup=never --changed-only`.
- `push --no-save` is valid only with `--session`.
- `run --direct` is valid only without `--arg` and without `--trace`; it uses `Excel.Run` directly and returns weaker diagnostics.
- `run --fast` uses direct execution when eligible and otherwise falls back to the normal temporary harness.
- `session start` keeps the configured workbook open and writes `.xlflow/session.json`.
- `--session` commands attach to the already-open workbook by path.
- `session stop` saves, closes, quits Excel, and removes session metadata.

## Outputs

- `push` may include top-level `source.changed`, `source.changed_only`, and `source.state`.
- `run.macro.direct` indicates whether the direct path was used.
- `session` commands may include top-level `session`.
- `runner` commands may include top-level `runner`.

# xlflow Diagnostic Run Spec

## Goal

Add an opt-in diagnostic run mode that converts VBE compile dialogs and VBA runtime failures into structured CLI output for AI agents.

## CLI Contract

- `xlflow run [macro] --diagnostic [--session] [--trace] [--headless | --interactive] [--timeout <duration>]`
- `--diagnostic --direct` is invalid.
- `--diagnostic --fast` is valid, but disables the fast direct path and uses the temporary harness.

## Behavior

- Diagnostic runs keep existing source preflight and GUI-boundary behavior.
- Before harness injection and macro invocation, diagnostic runs execute VBE Compile for the workbook VBA project.
- If VBE raises a modal compile dialog, xlflow reads the dialog text, closes the dialog, reads `ActiveCodePane.GetSelection`, and returns without invoking the macro.
- Runtime failures keep the existing harness-based `macro_failed` result and add `run_diagnostic.kind = "runtime"` when enriched diagnostics are available.
- v1 does not inject line numbers. Runtime line is reported only when VBA `Erl` is non-zero.

## Outputs

- Compile failure returns `error.code = "vba_compile_failed"`, `error.phase = "compile_vba"`, validation exit code `1`, and top-level `run_diagnostic`.
- Compile `run_diagnostic` includes `kind`, `message`, `location`, and `nearby_code` when VBE exposes them.
- Runtime `run_diagnostic` includes `kind = "runtime"` plus existing likely cause, suggestion, nearby source, and trace context fields.

# xlflow Session-Aware Defaults Spec

## Goal

Tighten the normal session workflow so agents can discover the active session automatically, surface save-required state clearly, and inspect richer runtime/build metadata from the CLI.

## CLI Contract

- `xlflow version [--verbose]`
- `xlflow save [--session]`
- `xlflow run [macro] [--session]`
- `xlflow pull|push|macros|test|trace ... [--session]`

## Behavior

- `version --verbose` includes the executable path, Go/build settings, PowerShell script resolution details, and a supported-feature list.
- When `run` is called without a macro argument, it uses `project.entry` from `xlflow.toml`.
- For `pull`, `push`, `macros`, `run`, `test`, `trace`, and `save`, omitting `--session` still reuses the matching live session workbook when `.xlflow/session.json` points at the configured workbook and the session is still open.
- Explicit `--session` and implicit auto-reuse are both preserved in result payloads via workbook session metadata.
- When a session-backed command leaves the live workbook newer than disk, the result includes structured `workbook.needs_save = true` and human output must make the save requirement obvious.
- `session status` includes whether the managed workbook is dirty and whether saving is currently required.
- `push` still saves by default; `--no-save` remains the session-only opt-out.

## Outputs

- Verbose version output may include `version.executable_path`, `version.go_version`, `version.module_path`, `version.build_settings`, `version.scripts`, and `version.features`.
- Workbook-backed commands may include `workbook.session_mode = explicit|auto|managed|none`, plus `workbook.session_requested`, `workbook.auto_session`, and `workbook.needs_save`.
- `session status` may include top-level `session.dirty` and `session.needs_save`.

# Release Artifact Trust Hardening Spec

## Goal

Strengthen trust signals for GitHub Release artifacts without changing the current Windows-first packaging model.

## Release Contract

- GitHub Releases continue to publish `xlflow_windows_x86_64.zip`.
- Releases publish a stable top-level checksum file named `checksums.txt`.
- `checksums.txt` uses SHA256.
- Releases publish SBOM files generated from archive artifacts.
- Release workflow publishes GitHub artifact attestations for release archives, `checksums.txt`, and generated SBOM artifacts.

## User Verification Contract

- Integrity verification: compare the ZIP SHA256 against `checksums.txt`.
- Provenance verification: run `gh attestation verify <zip> --repo harumiWeb/xlflow`.
- Non-claim: neither verification step is documented as proof of Windows publisher identity or Authenticode signing.

# xlflow Style-Aware Inspect Spec

## Goal

Extend file-based `inspect` so agents can verify worksheet output that depends on styling, layout, and merged-cell structure in addition to cell values.

## CLI Contract

- `xlflow inspect range [<sheet!A1:B2>] [--sheet <name> --address <A1:B2>] [--max-rows <n>] [--max-cols <n>] [--include-style] [--format text|json|markdown]`
- `xlflow inspect used-range [<sheet>] [--sheet <name>] [--max-rows <n>] [--max-cols <n>] [--include-style] [--format text|json|markdown]`

## Function and Type Contracts

```go
type TargetInfo struct {
    Kind string
    Path string
    Note string
}

type Limits struct {
    MaxRows int
    MaxCols int
}

type RangeOptions struct {
    Limits       Limits
    IncludeStyle bool
}

func SavedFileTargetInfo(path string) *TargetInfo
func Range(path, sheet, address string, opts RangeOptions) (RangeSnapshot, error)
func UsedRange(path, sheet string, opts RangeOptions) (RangeSnapshot, error)
```

- `SavedFileTargetInfo` returns the inspect target descriptor for saved workbook snapshots.
- `Range` accepts a workbook path, sheet name, A1 range selector, and bounded style options; it returns a file-based `RangeSnapshot`.
- `UsedRange` accepts a workbook path, sheet name, and bounded style options; it returns the lightweight used-range snapshot for the saved workbook file.

```go
type RangeSnapshot struct {
    Sheet         string
    Range         string
    UsedRange     string
    ReturnedRange string
    RowCount      int
    ColumnCount   int
    Values        [][]any
    Truncated     bool
    MaxRows       int
    MaxCols       int
    Warnings      []string
    StyleIncluded bool
    Cells         []StyledCellSnapshot
    Columns       []ColumnSnapshot
    Rows          []RowSnapshot
    MergedRanges  []string
}

type StyledCellSnapshot struct {
    Address             string
    Row                 int
    Column              int
    Value               any
    Formula             *string
    Fill                *CellFillSnapshot
    Font                *CellFontSnapshot
    Border              CellBorderSnapshot
    NumberFormat        *string
    HorizontalAlignment *string
    VerticalAlignment   *string
    Merged              bool
    MergeRange          *string
}
```

- `RangeSnapshot.Values` remains the compatibility matrix for existing inspect consumers.
- `RangeSnapshot.StyleIncluded` is omitted unless `RangeOptions.IncludeStyle` is true; when present it is always `true`.
- When `RangeOptions.IncludeStyle` is true, `Cells`, `Rows`, `Columns`, and `MergedRanges` are always present in JSON; empty selections serialize them as empty arrays.
- `StyledCellSnapshot` represents one returned cell in the truncated/output range, not the full unbounded workbook range.

## Behavior

- `inspect` continues to read the configured saved workbook file directly without starting Excel COM.
- `--include-style` is opt-in and only affects range-based inspect commands.
- Existing `values` matrix output remains unchanged for compatibility.
- When `--include-style` is set, range output also includes per-cell style metadata, row metadata, column metadata, and merged-range metadata for the returned range after truncation.
- Style-aware inspect reports the target as the saved workbook file and includes a note that unsaved live-session state is not being inspected.
- Empty cells are still included in the returned range and may carry style metadata even when `value` is `null`.
- Conditional formatting evaluation is out of scope for v1; output reflects stored cell styles in the saved workbook file.

## Outputs

- `inspect.target` remains the command target string such as `range` or `used-range`.
- `inspect.target_info` may include `kind`, `path`, and `note`.
- When `--include-style` is set, range snapshots deterministically include `style_included=true`, `cells`, `rows`, `columns`, and `merged_ranges`.
- When `--include-style` is absent, `style_included`, `cells`, `rows`, `columns`, and `merged_ranges` are omitted from JSON.
- Style cell objects include `address`, `row`, `column`, `value`, `formula`, `fill`, `font`, `border`, `number_format`, `horizontal_alignment`, `vertical_alignment`, `merged`, and `merge_range`.

# xlflow Range Image Export Spec

## Goal

Export a worksheet range as a PNG image through Excel COM so agents and users can visually verify workbook output that depends on layout, formatting, charts, fills, and sizing.

## CLI Contract

- `xlflow export-image [workbook] --sheet <name> --range <A1:B2> [--out <path> | --output-dir <dir>] [--name <filename>] [--format png] [--overwrite] [--session]`

## Function and Type Contracts

```go
type ExportImageOptions struct {
    WorkbookPath string
    Sheet        string
    Range        string
    OutPath      string
    OutputDir    string
    Name         string
    Format       string
    Overwrite    bool
    Session      bool
    Keepalive    CommandOptions
}

type ExportImageResolvedOutput struct {
    Path                 string
    Format               string
    Default              bool
    CreatedParentDirs    bool
}

func (r Runner) ExportImage(cfg config.Config, opts ExportImageOptions) (output.Envelope, int, error)
```

- `WorkbookPath` overrides `excel.path` for a one-off export target.
- `Range` must be validated and normalized to an uppercased A1 range before Excel opens when possible.
- `OutPath`, `OutputDir`, and `Name` determine the final output file path; `--out` takes precedence.
- `Format` is PNG-only in v1 and must agree with the chosen file extension when an extension is present.

## Behavior

- `export-image` follows the existing workbook-backed command behavior for sessions: when `.xlflow/session.json` points at the same workbook, omitting `--session` still auto-reuses that matching live session workbook.
- `--session` remains the explicit assertion mode and should fail if the matching session workbook is not available.
- Without `--out` or `--output-dir`, the default output directory is `.xlflow/artifacts/images/<workbook-name>/`.
- The default generated filename format is `<sheet>_<range>_<timestamp>.png`, with filesystem-safe sheet and range components.
- `--output-dir` chooses the directory while xlflow still generates the filename unless `--name` is also provided.
- `--name` chooses only the filename and uses the default workbook image directory unless `--output-dir` is also provided.
- Parent directories for the final output path are created automatically when missing.
- Existing output files fail with `output_file_exists` unless `--overwrite` is set.
- The implementation uses Excel COM `Range.CopyPicture`, pastes into a temporary `ChartObject`, exports via `Chart.Export`, then removes the temporary chart object.
- The temporary chart object must be removed on both success and failure paths when it has been created; cleanup failures after a successful export should surface as warnings rather than turning the command into a failure.
- For session-backed exports, result payloads should preserve the existing workbook save-state metadata such as `workbook.session_mode`, `workbook.dirty`, and `workbook.needs_save`.

## Outputs

- JSON output uses the existing top-level xlflow envelope fields `status`, `command`, `error`, and `logs`.
- `export-image` adds top-level `target`, `output`, and optional `warnings`.
- `target` includes the export source kind and workbook path; for session-backed results it may indicate `live_session`.
- `output` includes `path`, `format`, `default`, optional `created_parent_dirs`, and image dimensions when available.
- Human output should show the export target, workbook session/save-required state, output path, format, and dimensions when available.

## Error Contract

- Missing worksheet: `sheet_not_found`
- Invalid range selector: `invalid_range`
- Existing output without overwrite: `output_file_exists`
- Unsupported format or mismatched extension: `unsupported_image_format`
- Temporary chart cleanup after successful export: warning code `temporary_object_cleanup_failed`

## Explicit Workbook State Output

- Workbook-backed commands must make their read/write target explicit through top-level `target.kind = source|file|live_session`.
- Relevant commands must also return top-level `session.active`, `session.workbook_path`, `session.dirty`, and `session.save_required` when that state is knowable.
- Existing `workbook.session_mode`, `workbook.dirty`, and `workbook.needs_save` remain for compatibility; the new `target` and `session` blocks are the preferred stable contract for callers.
- `macros` with zero results should emit actionable `hints` explaining that discovery reads workbook state, not source files, and may require `push`.
- `inspect` keeps reading the saved workbook file directly, but when `.xlflow/session.json` points at the same live workbook it should surface session dirty state and warn when the saved file may be stale.

# xlflow Workbook Edit Commands Spec

## Goal

Add a minimal, session-only workbook mutation surface for AI-agent development loops so agents can prepare workbook state, trigger event-driven behavior, and tune layout without generating temporary VBA for every experiment.

## CLI Contract

- `xlflow edit cell [workbook] --sheet <name> --cell <A1> (--value <text> | --formula <formula> | --fill <#RGB|#RRGGBB>) --session [--events keep|on|off] [--keepalive] [--keepalive-interval <duration>]`
- `xlflow edit range [workbook] --sheet <name> --range <A1:B2> (--fill <#RGB|#RRGGBB> | --clear contents|formats|all) --session [--keepalive] [--keepalive-interval <duration>]`
- `xlflow edit rows [workbook] --sheet <name> --rows <1:31> --height <points> --session [--keepalive] [--keepalive-interval <duration>]`
- `xlflow edit columns [workbook] --sheet <name> --columns <A:AE> --width <chars> --session [--keepalive] [--keepalive-interval <duration>]`

## Function and Type Contracts

```go
type EditEventMode string

const (
    EditEventKeep EditEventMode = "keep"
    EditEventOn   EditEventMode = "on"
    EditEventOff  EditEventMode = "off"
)

type EditCellOptions struct {
    WorkbookPath string
    Sheet        string
    Cell         string
    Value        *string
    Formula      *string
    Fill         string
    Events       EditEventMode
    Session      bool
    Keepalive    CommandOptions
}

type EditRangeOptions struct {
    WorkbookPath string
    Sheet        string
    Range        string
    Fill         string
    Clear        string
    Session      bool
    Keepalive    CommandOptions
}

type EditRowsOptions struct {
    WorkbookPath string
    Sheet        string
    Rows         string
    Height       float64
    Session      bool
    Keepalive    CommandOptions
}

type EditColumnsOptions struct {
    WorkbookPath string
    Sheet        string
    Columns      string
    Width        float64
    Session      bool
    Keepalive    CommandOptions
}
```

## Behavior

- MVP requires `--session` for every `edit` command. xlflow must fail rather than silently editing a saved file or opening a hidden isolated workbook.
- `edit cell` supports exactly one mutation: value, formula, or fill.
- `edit range` supports exactly one mutation: fill or clear mode.
- `edit rows` supports row height only.
- `edit columns` supports column width only.
- Supported fill input formats are `#RGB` and `#RRGGBB`; normalized output uses uppercase `#RRGGBB`.
- `--events keep|on|off` applies only to `edit cell --value` and `edit cell --formula`.
- When `--events on` or `--events off` is used, xlflow captures the current `Application.EnableEvents` state, applies the requested temporary state, performs the edit, and restores the prior state whenever possible.
- Successful edits against a live session workbook must report dirty/save-required state and tell callers to use `xlflow save --session`.

## Outputs

- JSON keeps the existing xlflow top-level envelope.
- `edit` adds top-level `edit` metadata with `kind`, target selector, mutation summary, and optional event-state details.
- `target.kind` is always `live_session` in the MVP.
- `session.active=true`, `session.dirty=true`, and `session.save_required=true` after any successful edit.
- Human output must clearly report the edit target, mutation summary, and save-required state.
