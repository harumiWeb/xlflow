# xlflow `pack` Command (Experimental)

## Status

Experimental. This spec defines the contract for the pure-Go `pack` path, now implemented behind the required `--experimental` flag. When `pack` leaves experimental status, this contract folds into `docs/specs/cli-contract.md`. See `docs/adr/ADR-0012-pack-command.md` for the rationale and the `pack`/`push` boundary.

## Scope

`pack` builds a macro-enabled workbook artifact (`.xlsm`) from the xlflow source tree, entirely in Go at the file level. Template mode regenerates `xl/vbaProject.bin` and replaces that entry inside a workbook zip. Blank mode creates a fresh one-sheet OOXML workbook and VBA project without a workbook template. It never opens Excel and never uses COM or VBIDE, and it performs no VBE compile or runtime validation. `.xlsb` templates are rejected with `workbook_format_unsupported` because template mode expects an OOXML ZIP package.

`pack` does not change `push`. `push` remains the Excel/VBIDE-backed live-session path on Windows.

## Command

```text
xlflow [--json] pack --out <path.xlsm> [--template <path.xlsm> | --blank] --experimental
```

- `--out <path>` (required): destination artifact path. Must end in `.xlsm`. May overwrite an existing output file, but must not resolve to the template or the configured source workbook (see Template handling). A missing `--out` fails with `pack_args_invalid` (exit 2).
- `--template <path>` (optional): the workbook template the artifact is based on. When omitted, `pack` uses the source workbook configured in `xlflow.toml` under `[excel].path`. The template provides workbook structure — sheets, document-module hosts, and any existing designer streams; `pack` replaces only `xl/vbaProject.bin`.
- `--blank` (optional): create a fresh workbook without reading a template. It is mutually exclusive with an explicit `--template`; passing `--blank` together with `--template` — including an explicitly empty `--template ''` — fails with `pack_args_invalid` (exit 2). When neither flag is provided, template mode keeps the `[excel].path` fallback.
- `--experimental` (required while experimental): without it, `pack` fails with `pack_experimental_required` (exit 2).
- `--json`: persistent global flag; emits the standard envelope (see Output / JSON contract).

`--bridge` does not apply to `pack`. `pack` uses no Excel bridge.

## Template and source workbook handling

- The template is read-only. `pack` never writes back into it. `--out` must resolve to a different path than both the template and the configured source workbook; otherwise `pack` fails with `pack_in_place_overwrite` (exit 2).
- The identity comparison is canonical, not lexical. `pack` resolves the nearest existing ancestor of the template, the configured workbook, and the destination through symlinks and junctions (and normalizes Windows drive-letter case, short names, and UNC/extended prefixes) before comparing. A destination that names the same file through an aliased directory or a different lexical form is rejected with `pack_in_place_overwrite` even when it does not exist yet.
- `pack` operates only on closed workbook files. If a live xlflow session or an open workbook for the target is detected, `pack` fails with `pack_active_session` (exit 2) rather than reading possibly-dirty live state. The check covers the workbooks `pack` actually reads or replaces — the resolved template and the `--out` destination. When `--template` is omitted, the configured `[excel].path` workbook is the resolved template and is therefore covered. A session or lock file on any other workbook — including the configured workbook during a `--blank` or explicit `--template` pack — does not block packing; the in-place alias guard still applies to `--out`.
- Document-module hosts (`ThisWorkbook`, sheet modules) come from the template. `pack` maps document-module source onto them only when the mapping is unambiguous.

## Blank workbook profile

Blank mode has a fixed, fail-loud v1 topology: one workbook document module named `ThisWorkbook` and one worksheet/document module named `Sheet1`. Both `src/workbook/ThisWorkbook.bas` and `src/workbook/Sheet1.bas` must be present. Any additional or differently named document module fails with `pack_ambiguous_layout`. Any UserForm input fails with `pack_blank_userform_unsupported`; blank mode does not author `.frm`/`.frx` designer state. UserForm input is detected before full source-layout validation — `.frm`/`.frx` files anywhere under `[src].forms` and any file under its reserved `code/` or `specs/` directories all count — so malformed form artifacts (for example an orphan `.frx` or an unmatched sidecar) still report `pack_blank_userform_unsupported` rather than `pack_ambiguous_layout`.

The fresh project is named `VBAProject`, uses source-only module streams, and structurally authors the `PROJECT`, `PROJECTwm`, `dir`, `_VBA_PROJECT`, module, stdole, and Office reference records. The workbook and worksheet OOXML code names are `ThisWorkbook` and `Sheet1`, matching the VBA component identities. The output still uses atomic publication and must not alias the configured workbook.

`[pack.blank].code_page` and `[pack.blank].lcid` control locale-sensitive project records. Defaults are code page `1252` and LCID `1033`; for Japanese projects use `932` and `1041`. Supported code pages are the same set as template mode. Values are project metadata, not source-file encodings: managed source remains UTF-8 without BOM.

### VBA container compatibility

The `xl/vbaProject.bin` reader and writer support CFB major versions 3 and 4.
`pack` preserves the template container's major version: v3 templates remain
v3 with 512-byte sectors, and v4 templates remain v4 with 4096-byte sectors.
The writer emits DIFAT sectors when the FAT no longer fits in the 109 header
entries, so large carried-through streams do not impose the former header-only
FAT limit.

CFB sector chains, allocation tables, and directory references are validated
for bounds, cycles, duplicate ownership, declared length, and version geometry.
MS-OVBA compressed streams are limited to 64 MiB of decompressed data per
stream, and malformed compressed chunks or truncated `dir` records are
rejected. These failures are reported through `pack_ambiguous_layout`; `pack`
does not publish a partial artifact.

## Atomic publication

`pack` never writes the destination directly. Publication goes through a temporary sibling artifact in the destination directory (same volume):

1. Resolve the destination through existing symlinks and junctions. On Windows, when it already exists, probe it with exclusive sharing so a workbook that is locked or open elsewhere fails before staging work. Unix replacement relies on directory permissions and the final rename and therefore does not require the old file itself to be writable.
2. Write, flush, and close a uniquely named temporary artifact in the destination directory. New outputs use mode `0644` subject to the process umask; replacements retain the existing destination's permission bits.
3. Perform lightweight structural validation on the closed temporary artifact: it must be a readable OOXML zip with a non-empty, CRC-valid `xl/vbaProject.bin` entry.
4. Publish atomically. When no destination exists the temporary artifact is moved with a no-clobber atomic create (`publication="atomic_create"`); when one exists it is installed with the platform atomic replace API (`publication="atomic_replace"`). There is no delete-then-copy fallback.
5. The temporary artifact is removed on every failure path and after publication.

If generation, validation, or publication fails, an existing destination is left byte-for-byte unchanged. A destination that is locked or cannot be replaced safely fails with `pack_output_busy` (exit 3); a publication that cannot be performed atomically fails with `pack_output_replace_failed` (exit 3); staging or structural-validation failures remain `pack_write_failed` (exit 3). A cleanup failure before publication is joined to and returned with the primary failure. When temporary cleanup fails after a successful publication the command still succeeds, reports `output.temporary_cleanup.status="failed"` with `residual_path`, and emits a `pack_temporary_cleanup_failed` warning.

## Supported source (MVP)

- **Standard modules** (`.bas`) and **class modules** (`.cls`): the source tree is authoritative. Modules absent from the template are added, modules absent from source are removed, and renames are represented as removal of the old component plus addition of the new component. The textual `PROJECT` stream, `dir/PROJECTMODULES`, and `VBA/<stream>` entries are regenerated as one consistent component set.
- **Document modules**: supported only where they map safely against the template's existing document modules.
- **UserForm code-behind**: a form already present in the template has its code-behind updated from source, honoring `[userform].code_source`. In `frm` mode the code is read from `src/forms/*.frm`; in `sidecar` mode (the default) the authoritative code-behind is `src/forms/code/<FormName>.bas`, merged into the form in memory (the on-disk `.frm`/`.bas` are never modified — `pack` does not write sources). `src/forms/code` is a flat reserved directory; sidecar subdirectories are unsupported. In both modes only the code-behind is applied; the form's designer storage is carried through byte-for-byte. A matching `.frx` is inventoried and validated as a related source artifact but is not written into the packed project; `pack` never authors or modifies form layout. A `.frm` whose form is not in the template fails with `pack_userform_generation_unsupported`; a sidecar carrying `Attribute VB_*` header lines, using a subdirectory, or with no matching `.frm`, fails with `pack_ambiguous_layout`.
- **Existing UserForm designer streams in the template**: carried through byte-for-byte, untouched. `pack` does not generate or modify form layout.

In template mode, document-module and UserForm topology remains template-authoritative. Omitting their source does not remove them; source may update code only when the component already exists in the template. A source-only document module is rejected as `pack_ambiguous_layout`, and a source-only UserForm is rejected as `pack_userform_generation_unsupported`. Blank mode instead follows the fixed document topology above and rejects every UserForm.

## Code pages and component names

In template mode, the template's `dir` stream declares a `PROJECTCODEPAGE`; in blank mode, `[pack.blank].code_page` supplies it. Every MBCS text field in the project — module names, stream names, module doc strings, reference names, and the textual `PROJECT` stream — is decoded and re-encoded with that code page. Supported values are the practically relevant Windows ANSI code pages `874`, `932`, `936`, `949`, `950`, `1250`–`1258`, and `65001` (UTF-8). Any other code page fails deterministically with `pack_ambiguous_layout`.

Component names may contain any character representable in the project code page (for example, Japanese module names in a CP932 project). `pack` writes the paired MS-OVBA records (`MODULENAME`/`MODULENAMEUNICODE`, `MODULESTREAMNAME` and its embedded Unicode form, `MODULEDOCSTRING` and its embedded Unicode form) so the MBCS and UTF-16 forms always carry identical text. A name that the project code page cannot represent, a name containing NUL or line breaks, or a stream name longer than 31 UTF-16 code units fails loudly before any artifact is produced.

Stream names are directory entries in the CFB container, so their uniqueness follows the [MS-CFB] case-insensitive comparison (per-UTF-16-unit uppercase plus code-unit length), not Unicode lowercase — in a CP1254 project dotted `I` and dotless `ı` collide and are rejected at planning and at write time. A stream name matching a reserved VBA storage stream (`dir`, `_VBA_PROJECT`, compared with the same rule) is also rejected. The textual `PROJECT` stream never receives a component name containing line breaks, NUL, a double quote, or surrounding whitespace, and module records with IDs the writer owns cannot be supplied as opaque extras.

## Module metadata

When regenerating `dir/PROJECTMODULES`, `pack` preserves the module-level metadata of existing components: doc string, help context, read-only flag, private flag, module type, stream name, and any module records the writer does not interpret (re-emitted verbatim before the module terminator). New components added from source start with an empty doc string, help context `0`, no flags, and the MODULETYPE implied by their source type. `MODULECOOKIE` is ignored on read and always written as `0xFFFF`, and `MODULEOFFSET` is always written as `0` because packed projects are source-only — both follow the MS-OVBA write contract rather than preserving template values.

## Read-only source plan

Before building the source plan, `pack` validates the complete managed VBA source scope as UTF-8 without BOM. The scan covers configured module, class, form, and workbook roots plus the legacy top-level `tests/` root. Configured roots use the same separator normalization and absolute-path resolution as the pack source inventory, including absolute roots outside the project; external diagnostic paths remain absolute. This pack-only compatibility does not relax the project-containment rules for explicit `encoding check` or `encoding convert` paths. The scan includes `.frm` designer files even in sidecar mode and includes sidecar `.bas` files; binary `.frx` files are excluded. Invalid input returns `source_encoding_invalid` (exit 1) with the same source path, byte position, status, and remediation suggestions as `encoding check`. No source planning, template payload read, `vbaProject.bin` generation, or output publication occurs after this failure. `pack` never converts source implicitly; use `encoding convert --from cp932` explicitly for eligible CP932 input.

After encoding validation, `pack` builds a deterministic, read-only plan. Source discovery is shared with `build`, but `[build].exclude` is deliberately not applied to `pack`. Every configured source root must exist and be readable. Unknown extensions, invalid component names, case-insensitive duplicate names across component types, incomplete or ambiguous UserForm artifacts, and filename/`Attribute VB_Name` identity mismatches fail before project mutation or output publication. Template mode additionally rejects source-only document modules and source-only UserForms; blank mode applies its fixed topology contract.

The plan records each component's primary source path, related UserForm artifact paths, component type, topology authority, code authority, and action (`add`, `update`, `remove`, or `preserve`). Existing template order is retained. Source-authoritative standard/class additions are appended in stable type/name/path order. Standard/class topology and code are source-authoritative; document topology is template-authoritative while its supplied code is source-authoritative; UserForm topology/designer state is template-authoritative while its code follows `[userform].code_source`. Source files, the template, and the output artifact are not modified while the plan is built.

## Unsupported cases (fail-loud)

Each unsupported case is a specific, loud error. `pack` never falls back to best-effort behavior.

| Case                                                | Error code                                      | Exit |
| --------------------------------------------------- | ----------------------------------------------- | ---- |
| active xlflow session / live workbook               | `pack_active_session`                           | 2    |
| in-place overwrite of the template/source workbook  | `pack_in_place_overwrite`                       | 2    |
| aliased template/workbook path via symlink/junction | `pack_in_place_overwrite`                       | 2    |
| output locked or open in another process            | `pack_output_busy`                              | 3    |
| atomic publication impossible (no fallback)         | `pack_output_replace_failed`                    | 3    |
| protected VBA project                               | `pack_protected_project`                        | 1    |
| signed VBA project                                  | `pack_signed_project`                           | 1    |
| managed source is not UTF-8 without BOM             | `source_encoding_invalid`                       | 1    |
| creating a new UserForm / `.frx` generation         | `pack_userform_generation_unsupported`          | 1    |
| any UserForm in blank mode                          | `pack_blank_userform_unsupported`               | 1    |
| unknown or ambiguous VBA project layout             | `pack_ambiguous_layout`                         | 1    |
| missing `--out`, bad extension, other arg errors    | `pack_args_invalid`                             | 2    |
| missing `--experimental`                            | `pack_experimental_required`                    | 2    |
| template/source workbook not found or unreadable    | `pack_template_not_found`                       | 2    |
| source inventory or artifact staging failure        | `pack_source_read_failed` / `pack_write_failed` | 3    |
| `.xlsb` template selected by template mode          | `workbook_format_unsupported`                   | 2    |

## Output / JSON contract

On success with `--json`, `pack` emits the standard envelope (`status`, `command = "pack"`, `error = null`, `logs`) plus two top-level fields. An encoding failure instead uses the shared source-encoding failure envelope: `command = "pack"`, `error.code = "source_encoding_invalid"`, and `source.expected`, `source.files[]`, and `source.summary` describe the complete managed-source check.

- `output`: the produced artifact, mirroring the `export-image` `output` object — `path`, `format` (`"xlsm"`), optional `created_parent_dirs`, and the publication contract shared with `build`: `replaced_existing` (bool), `publication` (`"atomic_create"` or `"atomic_replace"`), and `temporary_cleanup` (`{"status": "clean"|"failed", "residual_path"?, "error"?}`).
- `pack`: identifies the backend and the validation posture.

```json
{
  "status": "ok",
  "command": "pack",
  "error": null,
  "output": {
    "path": "dist/Book.xlsm",
    "format": "xlsm",
    "replaced_existing": true,
    "publication": "atomic_replace",
    "temporary_cleanup": { "status": "clean" }
  },
  "pack": {
    "backend": "pure-go",
    "base": "template",
    "experimental": true,
    "vbe_validation": "not_performed",
    "template": "build/Book.xlsm",
    "modules": { "standard": 3, "class": 2, "document": 1, "form": 1, "carried_streams": 4 }
  },
  "warnings": [
    {
      "code": "vbe_validation_skipped",
      "message": "pack did not open Excel; no VBE compile or runtime validation was performed."
    }
  ],
  "logs": []
}
```

The backend identifier `pack.backend = "pure-go"` is deliberately distinct from the Excel-bridge `bridge` metadata defined in `cli-contract.md`, because `pack` uses no Excel bridge process. `pack.base` is `template` or `blank`; `pack.template` is present only for template mode. The `vbe_validation_skipped` warning is emitted on every successful run. Machine consumers must read `pack.vbe_validation` — not the absence of errors — to decide whether the artifact has been VBE-validated; it never is.

## Exit codes

`pack` follows the shared exit-code contract in `cli-contract.md`:

- `0`: success.
- `1`: validation/content failure detected from the project or template — protected project, signed project, ambiguous layout, unsupported UserForm generation.
- `2`: CLI argument or configuration error — missing `--out`, missing `--experimental`, bad extension, in-place overwrite, active session, template not found.
- `3`: environment failure — I/O failure reading sources or the template after validation has passed, staging the artifact, or publishing it (busy destination, non-atomic replace).

## No VBE validation contract

`pack` never compiles or executes VBA. The generated workbook is produced at the file level; Excel compiles it from source on first open. `pack` therefore cannot detect compile errors, runtime errors, or host-specific interpretation differences. This is a permanent property of the pure-Go path, not a temporary MVP gap. Every run reports `pack.vbe_validation = "not_performed"` and emits the `vbe_validation_skipped` warning. Consumers that need compile or runtime validation must use the Excel/VBIDE-backed `push` path on Windows.

## Linux test strategy

`pack` is a Go-core, file-level path, so it is fully testable on Linux without Excel:

- **Golden / byte-exact tests**: regenerate `xl/vbaProject.bin` for fixture projects and compare against committed golden bins; round-trip read → write → read for stability.
- **Source cross-checks**: decompress the regenerated module streams and compare against expected source text (and, where available, against `olevba` output) to confirm MS-OVBA correctness.
- **Negative tests**: each unsupported case asserts the documented error code and exit code.
- **Parser fuzzing**: native Go fuzz targets cover CFB parsing and round trips,
  MS-OVBA decompression and `dir` parsing, encrypted-data decoding, and complete
  `vbaProject.bin` reads. Successful CFB parses are rewritten and reparsed to
  check stream contents and major-version preservation.

These run on the existing Linux PR CI lane; no Windows runner is required for the `pack` path, and the automated PR path stays Linux/pure-Go only. Windows/Excel smoke tests that open the artifact and run a macro are a pre-stable milestone performed in the release gate, not in PR CI; see _Release-gate Excel smoke_ below.

## Release-gate Excel smoke (pre-stable)

The pure-Go path cannot tell whether a generated workbook actually compiles and runs (see _No VBE validation contract_). To close that gap before `pack` leaves experimental status, a representative packed artifact is smoke-tested in real Excel at the release gate — not in PR CI.

This smoke is part of the repository's manual release-gate flow (the `xlflow-tmp-workspace-e2e` skill, "pack artifact smoke" section); the automated PR path remains Linux/pure-Go only. The procedure:

1. Build a workspace with a known sentinel macro — a standard module that writes a fixed value to a cell.
2. Produce the artifact at the file level without template bytes: `xlflow pack --blank --out dist/Book.xlsm --experimental`. Confirm the JSON reports `pack.base = "blank"` and `pack.vbe_validation = "not_performed"`.
3. Open the produced `.xlsm` in Excel via COM, run the packed macro (which forces a VBE compile), and assert that the sentinel cell holds the expected value.

The smoke passes only if `pack` exits `0`, the workbook opens without a compile error, and the macro's observable effect (the sentinel cell) matches. A compile error, a wrong sentinel value, or a non-zero exit blocks the release.

This is release-gate evidence for a representative build. It does not change `pack`'s permanent `vbe_validation = "not_performed"` contract: `pack` itself never validates, and artifacts produced in the field remain unvalidated until opened in Excel.

## Staged UserForm plan

See ADR-0012 for the rationale. Summary:

- **Stage 1 (implemented)**: carry existing template designer streams through untouched; never generate forms.
- **Stage 2 (implemented)**: update the code-behind of a form already in the template, keeping the template's designer state. The code-behind source follows `[userform].code_source` — `frm` (code in the `.frm`) or `sidecar` (code in `src/forms/code/<FormName>.bas`, merged in memory); `pack` never writes the source tree. A `.frm` whose form is not in the template returns `pack_userform_generation_unsupported` (creating a form is Stage 3). A matching `.frx` is inventoried and validated as a related source artifact but is not written into the packed project.
- **Stage 3**: full reconstruction from exported `.frm`/`.frx` — a separate, higher-risk phase.

The reference implementation has demonstrated the Stage 1 / Stage 2 round-trip against real Excel, including a nested Frame/MultiPage form whose designer sub-storages round-trip byte-for-byte. That informs the staging but is not part of the MVP.
