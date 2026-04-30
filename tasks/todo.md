# xlflow VBA Source Encoding Todo

## Implementation

- [x] Add explicit UTF-8 without BOM and CP932 text helpers in the PowerShell bridge.
- [x] Convert exported VBIDE text from CP932 to UTF-8 during `pull`.
- [x] Convert UTF-8 source to CP932 temporary files during `push`.
- [x] Preserve `.frx` userform companion files as binary copies.
- [x] Keep workbook document module synchronization on UTF-8 source text.
- [x] Update CLI contract and ADR-0001 with the encoding boundary.

## Verification

- [x] Add helper tests for Japanese UTF-8 to CP932 round-trip behavior.
- [x] Add helper tests for byte-preserving `.frx` copy behavior.
- [x] Update document module tests to avoid default PowerShell text encoding.
- [x] Run `go test ./...` with a long timeout for Excel COM-backed tests.
- [x] Run `task verify`.
- [x] Run Excel COM e2e with Japanese VBA strings when available.
