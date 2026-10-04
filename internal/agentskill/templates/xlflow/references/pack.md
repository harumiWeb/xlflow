# Pack and Saved-File Artifact Reference

Load this reference when generating a separate `.xlsm`, choosing template versus
blank mode, or deciding whether an artifact has sufficient release evidence.
For FormSpec authoring, load [forms.md](forms.md).

## Choose the Output Boundary

`pack` is pure Go and cross-platform. It does not open Excel, compile VBA, or
execute macros. It creates a separate artifact and does not rewrite tracked
sources. Use file push for changes to the configured saved development workbook;
use Excel push for live import and VBE compile validation.

Use a template when workbook structure, sheets, styles, formulas, or existing
Designer state must survive. An explicit template avoids relying on the default
`[excel].path` fallback:

```bash
xlflow pack --template templates/Base.xlsm --out dist/Release.xlsm --json
```

Template and output must be closed and unowned by a live session. Output must
not alias the template or configured development workbook. Use a separate
destination instead of attempting an in-place pack. Read structured lock,
validation and publication errors; do not delete another process's lock.

## Blank Profile

Use blank mode when one fresh worksheet is sufficient and there is no workbook
template to preserve:

```bash
xlflow pack --blank --out dist/Release.xlsm --json
```

`--blank` and `--template` are mutually exclusive. The profile has exactly the
document modules ThisWorkbook and Sheet1, with both files required:

```text
src/workbook/ThisWorkbook.bas
src/workbook/Sheet1.bas
```

Each file can contain `Option Explicit` as its initial code. Additional or
differently named document modules fail. Keep all configured source roots
present/readable even when empty. Add standard/class modules in their source
roots and supported forms in `forms/specs/` with code sidecars as in [forms.md](forms.md).
Canonical form generation requires no `.frx`; the Forms reference is added when
needed. Blank mode does not reproduce a template's worksheet data or layout.

For an existing project using sidecar code and Japanese captions/names, configure:

```toml
[userform]
code_source = "sidecar"

[pack.blank]
code_page = 932
```

Blank code page defaults to 1252; template mode uses the template project's code
page. Managed VBA source is still UTF-8 without BOM. Root form captions and
component names must be representable in the project code page; control text
has its own Unicode persistence. Do not invent a `pack.blank.lcid` setting or
put unsupported configuration keys under `[pack]`.

## Module and Form Authority

| Source/template state                         | Template pack behavior                                                |
| --------------------------------------------- | --------------------------------------------------------------------- |
| Standard/class module supplied in source      | Add/update it; omission removes a template-only standard/class module |
| Document module supplied                      | Update matching template document code; new document topology fails   |
| Document module omitted                       | Preserve the template document module                                 |
| Canonical form spec supplied                  | Add a supported form or edit the existing Designer and code           |
| Form omitted with default topology            | Preserve the template form                                            |
| Existing legacy `.frm` without canonical spec | Code-only update; retain saved Designer                               |

For a canonical form in sidecar mode, code comes from matching `code/<Name>.bas`,
then matching `.frm` code, then empty code. An explicitly empty sidecar is
authoritative. In frm mode, matching `.frm` code is required. Pack does not use
file push's missing-sidecar-preserves-code rule. Keep explicit code sidecars
when packing file-pulled canonical specs so stale `.frm` code does not become a
fallback authority; a selected unsynchronized compatibility `.frm` fails FRM201.

To deliberately make the complete source form set authoritative, use:

```toml
[pack]
userform_topology = "source"
```

This removes template forms omitted from source, including the final form, and
requires canonical specs for supplied forms. Use it only when deletion is
intended. Default `"template"` preserves omitted forms. Blank mode always uses
source form topology; file push always preserves omitted saved forms regardless
of this setting. Within a supplied spec, its complete control collection is
authoritative while omitted retained properties preserve values.

## Supported Designer Work

Canonical specs support empty forms, common controls, nested Frames,
MultiPage/Page, standalone TabStrip tabs and BMP/JPEG Image assets. Template
pack can apply supported property/topology changes without regenerating unrelated
opaque state. Read [forms.md](forms.md) for hierarchy, pictures, omission semantics
and dimension limits. Custom ActiveX, unsupported persisted list state and
unsupported edits fail before publication instead of silently dropping bytes.

Do not treat `.frm`/`.frx` exports as sufficient input for a brand-new Designer.
New pure-Go forms require canonical specs; legacy existing forms can retain
their template Designer while only code changes.

## Validation and Release Evidence

Validate source with lint/analyze when relevant. Pack checks managed text
encoding, source/component identity, topology and binary readback before atomic
publication. Validated referenced picture bytes are assets rather than VBA text,
even if they have a source-like suffix. Pack does not apply `[build].exclude`;
clean the intended source set instead of assuming build exclusions affect pack.

Typical failures include `source_encoding_invalid`, `pack_ambiguous_layout`,
`pack_userform_generation_unsupported`, specific `pack_userform_edit_*` errors,
`pack_active_session` and `pack_output_busy`. Read the error and remediation;
do not strip Designer warnings, remove containers, or weaken source authority
just to get an artifact. A failed validation/publication leaves prior output
unchanged.

Successful JSON reports `pack.backend="pure-go"`, `pack.base="template"` or
`"blank"`, and `pack.vbe_validation="not_performed"`, with a
`vbe_validation_skipped` warning. Report artifact generation and runtime proof
separately. When compile/runtime acceptance is required, point an isolated
xlflow project at the generated artifact, start one Excel session on Windows,
run a focused sentinel/test, inspect the observable result, then save and stop.
Inspect forms visually and verify save/reopen when those are acceptance criteria.
Do not push unrelated development source over the artifact before verifying it.

For artifact-only CI without Excel, report structural generation success and
the unverified compile/runtime/visual scope. Do not open Excel or claim behavioral
success merely because packing succeeded.
