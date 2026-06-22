# ADR-0014: Deprecate the PowerShell Bridge

## Status

Accepted

## Context

xlflow originally used a bundled PowerShell bridge for Windows Excel automation. ADR-0008 introduced the `.NET` Excel bridge as the planned replacement, and ADR-0009 defined a provider contract that allowed staged migration through `auto` fallback to PowerShell while `.NET` reached parity.

The `.NET` bridge now covers the workbook-backed command set that previously required PowerShell, including `new`, `init`, `doctor`, `pull`, `push`, `run`, `test`, `save` / `session`, and related inspection and utility commands. It also owns diagnostics, dialog handling, structured errors, non-interactive behavior, UTF-8 transport, and session safety improvements that are increasingly expensive to duplicate in PowerShell.

Keeping both implementations active means every Excel COM, VBIDE, path/encoding, dialog, macro, and session edge case either has to be fixed twice or remains inconsistent across providers. That ambiguity makes support and release validation harder.

## Decision

Deprecate the legacy PowerShell bridge in v0.15.0 and remove it in v0.16.0.

In v0.15.0:

- Windows `auto` mode selects the `.NET` bridge and does not fall back to PowerShell.
- Explicit `.NET` selection remains strict.
- PowerShell remains available only through explicit opt-in: `--bridge powershell`, `XLFLOW_EXCEL_BRIDGE=powershell`, or `[excel].bridge = "powershell"`.
- Any PowerShell bridge selection emits a `powershell_bridge_deprecated` warning that states the v0.16.0 removal target.
- The PowerShell bridge is frozen except for critical safety or regression fixes.

In v0.16.0, remove PowerShell bridge selection, implementation, tests, documentation, and bundled bridge scripts.

This decision supersedes the fallback portions of ADR-0008 and ADR-0009. Their provider abstraction and `.NET` bridge process boundaries remain valid, but automatic PowerShell fallback is no longer part of the current contract.

## Consequences

- Positive: workbook-backed command behavior becomes easier to reason about because `auto` has one provider.
- Positive: missing `.NET` bridge executables, protocol mismatches, unsupported commands, invalid bridge JSON, and workbook-open failures are surfaced directly instead of being hidden by fallback.
- Positive: future Excel COM, VBIDE, dialog, and session fixes need only target the supported `.NET` bridge.
- Negative: users who still depend on PowerShell-specific behavior must opt in explicitly during v0.15.0 and report blockers before v0.16.0.
- Negative: environments that have not installed `xlflow-excel-bridge.exe` beside `xlflow.exe` will now fail in `auto` mode instead of silently using PowerShell.

## Alternatives Considered

1. **Keep PowerShell fallback through v0.15.0** - Rejected because fallback hides `.NET` bridge installation and compatibility problems during the deprecation window.
2. **Remove PowerShell immediately in v0.15.0** - Rejected to leave one release for users to identify and report remaining blockers.
3. **Keep PowerShell as a permanent explicit backend** - Rejected because it preserves the duplicate maintenance and support burden that the `.NET` bridge was introduced to eliminate.

## Related

- `docs/adr/ADR-0008-dotnet-excel-bridge.md`
- `docs/adr/ADR-0009-bridge-provider-contract.md`
- `docs/bridge/dotnet-bridge.md`
- `docs/bridge/powershell-bridge-legacy.md`
- `docs/specs/cli-contract.md`
- xlflow issue #177
