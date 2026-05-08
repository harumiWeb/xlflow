# xlflow Folder Structure Spec

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
- UserForm `.frm` and `.frx` companions must remain siblings after nested moves.
- Duplicate VBA component basenames anywhere in the recursive source tree fail `push` with `duplicate_module_name`.

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
