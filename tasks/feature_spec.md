# Headless File Dialog Wrapper Spec

## Goal

Extend the existing XlflowUI headless dialog path so file-selection flows can run unattended through explicit wrappers, while raw `Application.*` file dialog APIs remain blocked by GUI-boundary preflight.

## Contract

- New scaffolded `XlflowUI.bas` wrappers provide headless-compatible file dialog entrypoints instead of making raw `Application.GetOpenFilename`, `Application.GetSaveAsFilename`, or `Application.FileDialog(...)` directly headless-safe.
- Raw `Application.GetOpenFilename`, `Application.GetSaveAsFilename`, and `Application.FileDialog` remain `VB007` / `gui_boundary_detected` findings in headless runs.
- `xlflow run` and `xlflow test` accept a repeated `--filedialog <kind>:<dialog-id>=<value>` flag.
- Supported `kind` values are `get-open`, `file-open`, `save-as`, and `folder`.
- Dialog ids use the same normalization and collision rules as `--msgbox` / `--inputbox`.
- Repeating the same `kind:id` pair accumulates ordered selections for open-style dialogs.
- `@cancel` is the explicit scripted cancel token and cannot be combined with path values for the same dialog.
- `save-as` and `folder` reject multiple scripted values before Excel starts.
- `get-open` and `file-open` accept repeated scripted values at the CLI layer; wrapper calls that use `MultiSelect:=False` reject multiple scripted values when the VBA wrapper resolves the response.
- Headless UI events expose file-dialog results in structured `ui.events` output and realtime `--ui-stream` stderr summaries.

## VBA Contract

`XlflowUI.bas` adds these public wrappers:

```vb
Public Function GetOpenFilename(ByVal Id As String, Optional ByVal FileFilter As String = "", Optional ByVal FilterIndex As Long = 1, Optional ByVal Title As String = "", Optional ByVal ButtonText As String = "", Optional ByVal MultiSelect As Boolean = False, Optional ByVal DefaultValue As Variant) As Variant
Public Function FileDialogOpen(ByVal Id As String, Optional ByVal Title As String = "", Optional ByVal ButtonText As String = "", Optional ByVal MultiSelect As Boolean = False, Optional ByVal DefaultValue As Variant) As Variant
Public Function GetSaveAsFilename(ByVal Id As String, Optional ByVal InitialFileName As String = "", Optional ByVal FileFilter As String = "", Optional ByVal FilterIndex As Long = 1, Optional ByVal Title As String = "", Optional ByVal ButtonText As String = "", Optional ByVal DefaultValue As Variant) As Variant
Public Function FolderPicker(ByVal Id As String, Optional ByVal Title As String = "", Optional ByVal ButtonText As String = "", Optional ByVal InitialPath As String = "", Optional ByVal DefaultValue As Variant) As Variant
```

- All wrappers return `Variant` so they can preserve Excel-compatible interactive behavior while representing headless single-path, multi-path, and explicit-cancel results.
- In interactive mode, wrappers delegate to the corresponding native Excel API.
- In headless-like modes, wrappers resolve scripted values from xlflow-injected workbook markers keyed by normalized dialog id and kind.
- Explicit cancel returns `False`.
- Single-select wrappers return a `String` path.
- Multi-select wrappers return a one-dimensional `Variant` array of `String` paths in the order provided by CLI fixtures.
- Missing scripted responses remain deterministic XlflowUI-owned errors unless a wrapper-specific `DefaultValue` is supplied.

## CLI / Transport Contract

- `--filedialog` literals use `kind:id=value`.
- `value` is an opaque path string except for reserved `@cancel`.
- JSON transport between Go and PowerShell uses a dedicated `FileDialogResponsesJSON` payload, not the existing MsgBox/InputBox scalar maps.
- Each encoded dialog response includes `kind`, normalized `dialog_id`, ordered `values`, and `cancelled` state.
- Workbook defined names continue to be the runtime marker mechanism; the new payload only changes how file-dialog values are encoded and decoded.

## Output Contract

- `ui.events` for file dialogs may include `kind`, `dialog_id`, `response_source`, `cancelled`, `resolved_path`, `resolved_paths`, `error`, and `runtime_mode`.
- Human-readable UI summaries should keep single-select dialogs compact and summarize multi-select dialogs by count plus a short preview.

## Verification

- `go test ./internal/cli -run "FileDialog|UI|RunOptions|Test"`
- `go test ./internal/excel -run "FileDialog|UI|RunScriptArgs|TestScriptArgs"`
- `go test ./internal/excel/scripts -run "FileDialog|UI|RuntimeInjection"`
- `go test ./internal/project -run "FileDialog|UI"`
- `go test ./internal/gui ./internal/lint ./internal/output`
- Windows Excel COM E2E in `tmp_workspaces` covering single-select open, multi-select open, save-as, folder picker, and explicit cancel.

# winget Release Publishing Spec

## Goal

Publish xlflow Windows release metadata to winget using the existing `harumiWeb/winget-pkgs` fork and the `WINGET_GITHUB_TOKEN` release secret.

## Contract

- GitHub Releases continue to publish the existing Windows ZIP archive for `xlflow_windows_x86_64.zip`.
- GoReleaser generates a winget manifest for package identifier `HarumiWeb.Xlflow`.
- The winget manifest references the GitHub Release archive generated from the existing Windows archive configuration.
- GoReleaser pushes generated manifest changes to `harumiWeb/winget-pkgs` on branch `xlflow-<version>`.
- GoReleaser does not automatically open a cross-repository pull request to `microsoft/winget-pkgs`; upstream submission remains a manual follow-up from the pushed fork branch.
- Prerelease tags skip winget manifest publishing.
- The release workflow passes `WINGET_GITHUB_TOKEN` to GoReleaser.
- User-facing install docs include `winget install HarumiWeb.Xlflow` while keeping Scoop and GitHub Releases as supported alternatives.

## Verification

- `goreleaser check` validates the release configuration.
- `goreleaser release --snapshot --clean --skip=publish` validates archive, checksum, SBOM, and generated manifest behavior without publishing.
- `go test ./...` passes.
- After a real release, confirm the pushed branch and manifest path in `harumiWeb/winget-pkgs` before manually opening the upstream PR.

# Workbook Rollback Spec

## Goal

Add a first-class rollback flow for restoring the configured workbook from xlflow-managed backups after a broken `push`, failed import, or workbook-level mistake.

## Contract

- `xlflow backup list` lists rollback-capable workbook backups for the configured workbook.
- `xlflow rollback --latest` restores the newest matching backup.
- `xlflow rollback --backup <backup-id>` restores a specific matching backup.
- `push` default backup mode now creates workbook-file backups under `.xlflow/backups/<backup-id>/` with a copied workbook and `metadata.json`.
- Rollback creates a safety workbook backup with reason `pre-rollback` before replacing the target workbook file.
- Rollback restores only the workbook file; it does not update `src/` automatically.
- If the configured workbook is attached to an active xlflow session, rollback fails safely instead of replacing the file underneath the live workbook.
- Successful rollback warns that workbook and source may be out of sync and hints to run `inspect` and `pull`.

## Function / Type Contracts

```go
type Metadata struct {
    ID                   string
    CreatedAt            time.Time
    Reason               string
    OriginalWorkbookPath string
    BackupFilePath       string
}

type Record struct {
    Metadata
    Directory         string
    BackupFileAbsPath string
}

func List(projectRoot string, workbookPath string) ([]Record, error)
func Latest(projectRoot string, workbookPath string) (Record, error)
func Create(projectRoot string, workbookPath string, reason string, now time.Time) (Record, error)
func Restore(targetWorkbookPath string, record Record) error
```

- `List` filters by workbook path and ignores legacy backup directories that do not contain valid rollback metadata.
- `Latest` returns the newest matching backup record for the configured workbook.
- `Create` returns the created backup record, including the generated backup ID and resolved backup workbook path.
- `Restore` replaces the target workbook from a selected record and is paired with CLI-level `workbook_in_use` safety checks before file replacement.

## Verification

- CLI tests cover command registration and selector validation for `backup list` and `rollback`.
- Unit tests cover backup metadata listing, latest selection, workbook-path filtering, legacy directory ignore behavior, and restore behavior.
- `push.ps1` continues to parse and now creates workbook-file backup artifacts compatible with rollback.

# new/init Bootstrap Sync Spec

## Goal

Remove the extra bootstrap commands after project creation so `new` and `init` leave workbook state and source state synchronized immediately.

## Contract

- `xlflow new` still scaffolds `src/modules/XlflowAssert.bas`, `Main.bas`, `App.bas`, and `Ui.bas`, plus workbook document-module placeholders `src/workbook/ThisWorkbook.bas` and `src/workbook/Sheet1.bas`, then automatically `push`es that scaffolded source into the newly created workbook before reporting success.
- `xlflow init` still copies the input workbook into `build/<basename>`, then automatically `pull`s VBA from that copied workbook into `src/` before reporting success.
- `new` success output includes a log that scaffolded VBA was pushed to the workbook.
- `init` success output includes a log that workbook VBA was pulled into source.
- `init` is now an Excel COM-backed command for contract purposes: if the automatic bootstrap pull cannot complete because Excel, COM, PowerShell, or VBIDE access is unavailable, the command returns an environment failure instead of a partial success.
- `init` supports the same `--keepalive` / `--keepalive-interval` heartbeat behavior used by other Excel COM-backed commands.

## Verification

- `new` followed immediately by `pull` must preserve the scaffolded modules instead of exporting an empty workbook VBA project.
- `init` must materialize workbook VBA under `src/` without requiring a manual follow-up `pull`.
- CLI regression tests must cover the new success logs, `init` keepalive flags, and the bootstrap `push` / `pull` command chaining.

# VitePress Documentation Site Spec

## Goal

Publish an English-first xlflow documentation site from `vitepress/` with GitHub Pages project routing at `/xlflow/`.

## Contract

- VitePress uses `vitepress/.vitepress/config.mts` with local search, clean URLs, last-updated metadata, GitHub edit links, and sectioned navigation.
- The starter VitePress example pages are no longer linked from navigation.
- The site has stable pages for guides, concepts, command reference, AI-agent workflows, demos, reference material, and design background.
- Command pages must match the current Cobra command surface reported by `go run ./cmd/xlflow --help`.
- Public-facing prose is derived from the README files; stable CLI contracts and JSON/exit-code details are derived from `docs/specs/*`; design pages link back to ADRs.
- GitHub Pages deployment builds with pnpm and uploads `vitepress/.vitepress/dist`.
- If `vitepress/.vitepress/theme` exists, it must provide an `index` entrypoint. The default site theme is represented by re-exporting `vitepress/theme` from `vitepress/.vitepress/theme/index.ts`.

## Verification

- `pnpm docs:build` must pass.
- VitePress build output must not report broken internal links.
- Generated asset paths must be compatible with `base: "/xlflow/"`.

## VitePress Docs Polish Spec

## Goal

Improve first-run VitePress guidance for Windows users and AI-agent workflows without changing the CLI contract.

## Contract

- `vitepress/installation.md` presents the Windows end-user path first, with Scoop and GitHub Releases ahead of `go install`.
- `vitepress/quickstart.md` explains what `xlflow new` creates and where users should edit VBA source before `push`.
- `vitepress/ai-agents/index.md` documents a short recovery loop for failed agent runs and links the stable machine contract pages for JSON output and error codes.

## Verification

- `pnpm docs:build` must pass.
- Updated pages must keep existing internal links valid under the GitHub Pages base path.

# VBA Syntax Lint Spec

## Goal

Catch cheap VBA syntax mistakes before Excel/VBE can surface them as modal compile dialogs.

## Contract

- `xlflow lint` reports always-on `error` findings for malformed `Sub`, `Function`, and `Property Get/Let/Set` procedure boundaries.
- `VB010` reports an unterminated procedure at the procedure start line.
- `VB011` reports an unexpected `End Sub`, `End Function`, or `End Property` when no matching procedure is open.
- `VB012` reports a mismatched procedure end when the open procedure kind differs from the `End ...` kind.
- `VB013` reports a line-continuation underscore that is not preceded by whitespace.
- The scanner ignores comments and string literal contents for token detection and underscore checks, and joins valid continued physical lines before procedure boundary detection.
- These findings are push/run source-preflight blocking issues, matching existing compile-dialog prevention rules.

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

## Bundled Helper Module Install Spec

### Goal

Allow existing xlflow projects to adopt the bundled helper modules without recreating the project scaffold.

### Commands

- `xlflow init <workbook> --with-module`
- `xlflow module install [--push]`

### Behavior

- `init` keeps its existing bootstrap copy plus `pull` behavior when `--with-module` is absent.
- `init --with-module` installs `XlflowAssert.bas`, `XlflowRuntime.bas`, `XlflowUI.bas`, and `XlflowDebug.bas` after the bootstrap `pull`, then automatically `push`es them into the copied workbook so source and workbook remain synchronized.
- `module install` installs `XlflowAssert.bas`, `XlflowRuntime.bas`, `XlflowUI.bas`, and `XlflowDebug.bas` into the configured `[src].modules` root of an existing xlflow project.
- `module install` is source-only by default.
- `module install --push` reuses the normal push preflight and workbook import path after writing the helper source files.
- Both commands refuse to overwrite existing target helper source files.
- Existing-project helper installation must honor custom `[src].modules` roots instead of hardcoding `src/modules`.

### Verification

- `go test ./internal/project ./internal/cli`
- `xlflow init LegacyBook.xlsm --with-module`
- `xlflow module install`
- `xlflow module install --push`
- Collision case where one of `XlflowAssert.bas`, `XlflowRuntime.bas`, `XlflowUI.bas`, or `XlflowDebug.bas` already exists

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
- `inspect workbook|sheets|range|used-range|cell --session` reads the live workbook currently attached to the xlflow session through Excel COM and returns the same target-specific payload shapes with `inspect.source = "excel_com"` and `inspect.target_info.kind = "live_session"`.
- `--include-style` is opt-in and only affects range-based inspect commands.
- Existing `values` matrix output remains unchanged for compatibility.
- When `--include-style` is set, range output also includes per-cell style metadata, row metadata, column metadata, and merged-range metadata for the returned range after truncation.
- Style-aware inspect reports the target as the saved workbook file and includes a note that unsaved live-session state is not being inspected.
- File-backed inspect keeps warning when a matching live session is dirty and now adds hints to run the matching `inspect ... --session` command or `save --session`.
- Empty cells are still included in the returned range and may carry style metadata even when `value` is `null`.
- Conditional formatting evaluation is out of scope for v1; output reflects stored cell styles in the saved workbook file.

## Outputs

- `inspect.target` remains the command target string such as `range` or `used-range`.
- `inspect.target_info` may include `kind`, `path`, and `note`.
- When `--include-style` is set, range snapshots deterministically include `style_included=true`, `cells`, `rows`, `columns`, and `merged_ranges`.
- When `--include-style` is absent, `style_included`, `cells`, `rows`, `columns`, and `merged_ranges` are omitted from JSON.
- Style cell objects include `address`, `row`, `column`, `value`, `formula`, `fill`, `font`, `border`, `number_format`, `horizontal_alignment`, `vertical_alignment`, `merged`, and `merge_range`.

# Runtime-Aware Execution Mode Spec

## Goal

Expose a stable VBA-visible execution mode so workbook code can branch between normal Excel use and xlflow-managed automation without relying on process inspection hacks.

## Scope

- Phase 1 covers execution context resolution and injection for `xlflow run` and `xlflow test`.
- Phase 2 builds on that transport with `XlflowUI.MsgBox` and `XlflowUI.InputBox` plus repeated CLI flags `--msgbox <dialog-id=result>` and `--inputbox <dialog-id=value>` on both `run` and `test`.
- `XlflowUI` requires stable dialog ids. Ids must contain at least one ASCII letter or digit and are normalized to lowercase ASCII letters/digits separated by `_` for workbook-marker lookup.
- In `interactive` mode, `XlflowUI` delegates to native `VBA.Interaction.MsgBox` and `VBA.Interaction.InputBox`.
- In `headless`, `ci`, `agent`, and `test` modes, `XlflowUI` resolves scripted responses from xlflow-injected workbook markers and must fail deterministically when a required response is missing or invalid.
- Manual Excel usage remains backward compatible: when no xlflow marker is present, VBA falls back to interactive mode.

## Runtime Modes

- `interactive`
- `headless`
- `ci`
- `agent`
- `test`

- `IsHeadless` is true for `headless`, `ci`, `agent`, and `test`.

## Resolution Contract

- `xlflow run --headless` resolves to `headless`.
- `xlflow run --interactive` resolves to `interactive`.
- `xlflow test` resolves to `test`.
- When no explicit mode flag is set, xlflow may honor `XLFLOW_MODE=interactive|headless|ci|agent|test` as a wrapper-level override for unattended agent or CI entrypoints.
- Unknown `XLFLOW_MODE` values are ignored rather than guessed.
- If no xlflow-specific signal exists, VBA fallback remains `interactive`.

## Injection Contract

- Before user VBA starts, `run.ps1` and `test.ps1` write a workbook-scoped hidden defined name `__XLFLOW_MODE__` with the resolved mode string.
- xlflow may also write a reserved compatibility marker such as `__XLFLOW_RUNTIME_VERSION__ = 1` for future helper evolution.
- Phase 2 additionally writes temporary workbook-scoped hidden defined names for scripted `XlflowUI` dialog responses, keyed by normalized dialog id and wrapper kind (`msgbox` or `input`).
- Runtime markers are removed or restored in `finally`, even when macro/test execution fails.
- Cleanup occurs before `Save`, `SaveCopyAs`, and final result serialization so runtime markers are not persisted as workbook content.
- If a reserved name already exists, xlflow preserves its prior value and restores it after execution rather than silently overwriting it permanently.

## VBA API Contract

- New scaffolded projects add `src/modules/XlflowRuntime.bas`.
- Phase 1 uses VBA-native module-qualified calls rather than a pseudo-namespace. The stable helper surface is:

```vb
Private Const xlflowInteractive As Long = 0
Private Const xlflowHeadless As Long = 1
Private Const xlflowCI As Long = 2
Private Const xlflowAgent As Long = 3
Private Const xlflowTest As Long = 4

Public Function Mode() As Long
Public Function ModeName() As String
Public Function IsInteractive() As Boolean
Public Function IsHeadless() As Boolean
Public Function IsCI() As Boolean
Public Function IsAgent() As Boolean
Public Function IsTest() As Boolean
```

- `XlflowRuntime` reads workbook-scoped state first, then `Environ$("XLFLOW_MODE")`, then defaults to `interactive`.
- Phase 2 also scaffolds `src/modules/XlflowUI.bas` with the stable wrapper surface:

```vb
Public Function MsgBox(ByVal Id As String, ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "") As VbMsgBoxResult
Public Function InputBox(ByVal Id As String, ByVal Prompt As String, Optional ByVal Title As String = "", Optional ByVal Default As String = "") As String
```

- `XlflowUI` validates dialog ids before interactive or headless dispatch.
- `MsgBox` scripted responses accept `abort`, `cancel`, `ignore`, `no`, `ok`, `retry`, and `yes`.
- Existing projects are not auto-migrated in Phase 2; documentation should show how to add `XlflowRuntime.bas` and `XlflowUI.bas` manually when a project opts into runtime-aware dialogs.

## Output Contract

- `run` and `test` may include top-level `runtime.mode`, `runtime.mode_name`, `runtime.injected`, and `runtime.source = command|environment|default`.
- Human output should mention the resolved runtime mode when xlflow invokes user VBA.

## Implementation Notes

- Add a shared Go execution-mode resolver so `run`, `test`, docs, and regression tests use one mapping.
- Add shared PowerShell helpers in `common.ps1` for reserved-name add/restore and mode serialization.
- `New-XlflowRunHarnessCode` and the test runner keep their existing temporary-module flow; runtime marker injection is parallel to trace/helper injection rather than replacing it.
- `xlflow new` scaffolds `XlflowRuntime.bas`; `init` keeps existing workbook import behavior and does not force runtime helper migration into imported projects.

## Verification

- `go test ./internal/cli ./internal/excel ./internal/project`
- `go test ./internal/excel/scripts -run "TestPowerShellScriptsParse|TestBuildRunScriptArgsPassesRuntimeMode|TestRunScriptInjectsRuntimeMode|TestTestScriptInjectsRuntimeMode"`
- `go test ./...`
- Windows Excel COM validation in a disposable workspace:
  - `xlflow new --json`
  - `xlflow run Main.Run --interactive --json`
  - `xlflow run Main.Run --headless --json`
  - `XLFLOW_MODE=agent xlflow run Main.Run --json`
  - `xlflow test --json`
- Validation workbook code should write `XlflowRuntime.ModeName()` and the predicate booleans to worksheet cells so Excel state proves each injected mode reached VBA.

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
- Relevant commands must also return top-level `session.active`, `session.workbook_path`, `session.workbook_name`, `session.dirty`, and `session.save_required` when that state is knowable.
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
