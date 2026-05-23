# Coding Rules

## General

- Keep changes small, cohesive, and scoped to the requested behavior.
- Follow existing package boundaries, naming, and helper patterns.
- Prefer explicit, boring control flow over clever abstractions.
- Do not add speculative extension points, unused compatibility layers, or unrelated formatting churn.
- Validate user-visible file extensions against the actual serialized format.
- Apply defaults and normalization in one place, then treat normalized values as authoritative.
- Preserve specific structured errors; do not overwrite them with broad fallback errors.

## Go

- Keep generic tests cross-platform. Use `filepath.Join` instead of hardcoded separators.
- Gate shell-specific tests on executable availability with `t.Skip`.
- Search for unused wrappers or dead code after refactoring shared helpers.
- Keep CLI validation cheap when possible, especially before Excel COM or workbook opens.
- Maintain public JSON contracts and human output distinctions such as empty results versus unavailable results.

## PowerShell

- Shared scripts under `internal/excel/scripts` must remain compatible with Windows PowerShell 5.1.
- Avoid optional cmdlets for core bridge behavior; prefer stable .NET APIs available in Windows PowerShell 5.1.
- Wrap function calls in parentheses before combining them with `-and` or `-or`.
- Remember variable names are case-insensitive; do not parse string parameters into same-name local booleans with different casing.
- Prefer behavior tests that dot-source script helpers when pure decision logic changes.
- Use `finally` cleanup for temporary import/export workspaces and file handles.

## VBA Generation

- Generated VBA that opens files must track open state, close on errors, and re-raise the original error.
- Generated run harnesses should call possibly missing macros through workbook-qualified dynamic `Application.Run` and verify entrypoints.
- Temporary VBIDE helper modules must use collision-resistant names and avoid replacing user-visible modules.
- Do not compare or stringify object/array variants with generic assertion operators unless a precise contract exists.
- Normalize document-module exports to editable body text before linting or push.

## Excel Automation

- Prevent known VBE modal compile dialogs with source preflight before import or run.
- Keep automation macros executable for temp workbook copies that rely on injected helpers.
- Do not stop VBE compile dialog watchers immediately after compile execution; allow a real post-compile wait window.
- Prefer `xlflow run --diagnostic` for AI-agent workbook execution unless GUI dialogs are explicitly desired.

## Docs

- User-facing changes may require `CHANGELOG.md`, `README.md`, `vitepress/`, and `docs/specs/` updates.
- Generated documentation pages should be self-contained with real fenced commands, options/defaults, warnings, and JSON examples where useful.
