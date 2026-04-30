# xlflow VBA Source Encoding Spec

## Goal

Keep Git-managed VBA source files UTF-8 while preserving compatibility with Excel/VBIDE import and export behavior on Japanese Windows.

## Behavior

- Source-controlled `.bas`, `.cls`, and `.frm` files are UTF-8 without BOM.
- VBIDE import/export text files are treated as CP932 at the PowerShell bridge boundary.
- `xlflow pull` exports through VBIDE, reads exported text as CP932, and rewrites source files as UTF-8 without BOM.
- `xlflow push` reads source files as UTF-8 without BOM, writes CP932 temporary import copies under `.xlflow/tmp/import/<timestamp>/`, and imports those copies through VBIDE.
- `.frx` userform companion files are binary and are copied without text conversion.
- Workbook document modules are normalized and synchronized from UTF-8 source text.

## Verification

- PowerShell helper tests cover UTF-8/CP932 round-trip behavior for Japanese text.
- PowerShell helper tests cover byte-preserving `.frx` copy behavior.
- Document module normalization tests include Japanese body text and do not rely on default `Get-Content` or `Set-Content` encoding.
- Fast gate: `go test ./...` and `task verify`.
- Excel COM gate: run the `xlflow-tmp-workspace-e2e` workflow with Japanese VBA strings when Excel/VBIDE access is available.
