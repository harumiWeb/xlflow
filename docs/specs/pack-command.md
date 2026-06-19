# xlflow `pack` Command (Experimental)

## Status

Experimental. This spec defines the contract for the pure-Go `pack` path, now implemented behind the required `--experimental` flag. When `pack` leaves experimental status, this contract folds into `docs/specs/cli-contract.md`. See `docs/adr/ADR-0012-pack-command.md` for the rationale and the `pack`/`push` boundary.

## Scope

`pack` builds a macro-enabled workbook artifact (`.xlsm`) from the xlflow source tree plus a workbook template, entirely in Go at the file level. It regenerates `xl/vbaProject.bin` from `.bas`/`.cls` sources and replaces that single entry inside the workbook zip. It never opens Excel and never uses COM or VBIDE, and it performs no VBE compile or runtime validation.

`pack` does not change `push`. `push` remains the Excel/VBIDE-backed live-session path on Windows.

## Command

```text
xlflow [--json] pack --out <path.xlsm> [--template <path.xlsm>] --experimental
```

- `--out <path>` (required): destination artifact path. Must end in `.xlsm`. May overwrite an existing output file, but must not resolve to the template or the configured source workbook (see Template handling). A missing `--out` fails with `pack_args_invalid` (exit 2).
- `--template <path>` (optional): the workbook template the artifact is based on. When omitted, `pack` uses the source workbook configured in `xlflow.toml` under `[excel].path`. The template provides workbook structure — sheets, document-module hosts, and any existing designer streams; `pack` replaces only `xl/vbaProject.bin`.
- `--experimental` (required while experimental): without it, `pack` fails with `pack_experimental_required` (exit 2).
- `--json`: persistent global flag; emits the standard envelope (see Output / JSON contract).

`--bridge` does not apply to `pack`. `pack` uses no Excel bridge.

## Template and source workbook handling

- The template is read-only. `pack` never writes back into it. `--out` must resolve to a different path than both the template and the configured source workbook; otherwise `pack` fails with `pack_in_place_overwrite` (exit 2).
- `pack` operates only on closed workbook files. If a live xlflow session or an open workbook for the target is detected, `pack` fails with `pack_active_session` (exit 2) rather than reading possibly-dirty live state.
- Document-module hosts (`ThisWorkbook`, sheet modules) come from the template. `pack` maps document-module source onto them only when the mapping is unambiguous.

## Supported source (MVP)

- **Standard modules** (`.bas`) and **class modules** (`.cls`): regenerated into `xl/vbaProject.bin` from source.
- **Document modules**: supported only where they map safely against the template's existing document modules.
- **Existing UserForm designer streams in the template**: carried through byte-for-byte, untouched. `pack` does not generate or modify form layout.

## Unsupported cases (fail-loud)

Each unsupported case is a specific, loud error. `pack` never falls back to best-effort behavior.

| Case                                               | Error code                             | Exit |
| -------------------------------------------------- | -------------------------------------- | ---- |
| active xlflow session / live workbook              | `pack_active_session`                  | 2    |
| in-place overwrite of the template/source workbook | `pack_in_place_overwrite`              | 2    |
| protected VBA project                              | `pack_protected_project`               | 1    |
| signed VBA project                                 | `pack_signed_project`                  | 1    |
| full UserForm / `.frx` generation                  | `pack_userform_generation_unsupported` | 1    |
| unknown or ambiguous VBA project layout            | `pack_ambiguous_layout`                | 1    |
| missing `--out`, bad extension, other arg errors   | `pack_args_invalid`                    | 2    |
| missing `--experimental`                           | `pack_experimental_required`           | 2    |
| template/source workbook not found or unreadable   | `pack_template_not_found`              | 2    |

## Output / JSON contract

On success with `--json`, `pack` emits the standard envelope (`status`, `command = "pack"`, `error = null`, `logs`) plus two top-level fields:

- `output`: the produced artifact, mirroring the `export-image` `output` object — `path`, `format` (`"xlsm"`), and optional `created_parent_dirs`.
- `pack`: identifies the backend and the validation posture.

```json
{
  "status": "ok",
  "command": "pack",
  "error": null,
  "output": {
    "path": "dist/Book.xlsm",
    "format": "xlsm"
  },
  "pack": {
    "backend": "pure-go",
    "experimental": true,
    "vbe_validation": "not_performed",
    "template": "build/Book.xlsm",
    "modules": { "standard": 3, "class": 2, "document": 1, "carried_streams": 4 }
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

The backend identifier `pack.backend = "pure-go"` is deliberately distinct from the Excel-bridge `bridge` metadata defined in `cli-contract.md`, because `pack` uses no Excel bridge process. The `vbe_validation_skipped` warning is emitted on every successful run. Machine consumers must read `pack.vbe_validation` — not the absence of errors — to decide whether the artifact has been VBE-validated; it never is.

## Exit codes

`pack` follows the shared exit-code contract in `cli-contract.md`:

- `0`: success.
- `1`: validation/content failure detected from the project or template — protected project, signed project, ambiguous layout, unsupported UserForm generation.
- `2`: CLI argument or configuration error — missing `--out`, missing `--experimental`, bad extension, in-place overwrite, active session, template not found.
- `3`: environment failure — I/O failure writing the artifact or reading the template after validation has passed.

## No VBE validation contract

`pack` never compiles or executes VBA. The generated workbook is produced at the file level; Excel compiles it from source on first open. `pack` therefore cannot detect compile errors, runtime errors, or host-specific interpretation differences. This is a permanent property of the pure-Go path, not a temporary MVP gap. Every run reports `pack.vbe_validation = "not_performed"` and emits the `vbe_validation_skipped` warning. Consumers that need compile or runtime validation must use the Excel/VBIDE-backed `push` path on Windows.

## Linux test strategy

`pack` is a Go-core, file-level path, so it is fully testable on Linux without Excel:

- **Golden / byte-exact tests**: regenerate `xl/vbaProject.bin` for fixture projects and compare against committed golden bins; round-trip read → write → read for stability.
- **Source cross-checks**: decompress the regenerated module streams and compare against expected source text (and, where available, against `olevba` output) to confirm MS-OVBA correctness.
- **Negative tests**: each unsupported case asserts the documented error code and exit code.

These run on the existing Linux PR CI lane; no Windows runner is required for the `pack` path. Windows/Excel smoke tests that open the artifact and run a macro are a pre-stable milestone, not part of the MVP CI.

## Staged UserForm plan

See ADR-0012 for the rationale. Summary:

- **Stage 1 (MVP-compatible)**: carry existing template designer streams through untouched; never generate forms.
- **Stage 2**: update form code-behind while keeping the template's designer state.
- **Stage 3**: full reconstruction from exported `.frm`/`.frx` — a separate, higher-risk phase.

The reference implementation has demonstrated the Stage 1 / Stage 2 round-trip against real Excel, including a nested Frame/MultiPage form whose designer sub-storages round-trip byte-for-byte. That informs the staging but is not part of the MVP.
