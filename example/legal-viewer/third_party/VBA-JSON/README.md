# VBA-JSON

- Source: https://github.com/VBA-tools/VBA-JSON
- Vendored module: `src/modules/JsonConverter.bas`
- License: MIT, see `third_party/VBA-JSON/LICENSE`
- Version noted in source: VBA-JSON v2.3.1

## Local Compatibility Patches

- Replaced the `Dictionary` early-bound return/new expression with late binding via `CreateObject("Scripting.Dictionary")` to avoid requiring a manual `Microsoft Scripting Runtime` reference.
- Rewrote one quote-escape assignment as `Chr$(92) & Chr$(34)` so xlflow lint does not treat it as C-style escaping.
- Made `JsonOptions` private and removed broad `On Error Resume Next` blocks so the module passes this project's xlflow lint rules.
