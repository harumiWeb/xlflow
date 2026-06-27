# ADR-0009: Bridge Provider Contract and Fallback Rules

## Status

Accepted

## Context

ADR-0008 introduced a .NET bridge, but xlflow already had a working PowerShell bridge and many workbook-backed commands depended on established JSON envelope and exit-code behavior. A safe migration required a stable provider contract so commands could move to .NET incrementally while unsupported commands kept working through PowerShell in `auto` mode.

The contract must also let CI and users prove that a command did or did not use PowerShell. Silent fallback from an explicit `.NET` request would hide corporate execution-policy problems and weaken validation.

## Decision

Introduce a bridge provider abstraction with three user-facing modes:

- `auto`
- `dotnet`
- `powershell`

Selection priority:

```txt
CLI option --bridge
  > XLFLOW_EXCEL_BRIDGE
  > [excel].bridge in xlflow.toml
  > default auto
```

Fallback rules:

- `--bridge dotnet` is strict. If `xlflow-excel-bridge.exe` is missing, incompatible, or does not support the command, the command fails with a structured bridge error. It must not silently invoke PowerShell.
- `--bridge powershell` originally used the PowerShell bridge during migration. ADR-0014 later removed this mode.
- `--bridge auto` originally allowed fallback to PowerShell during migration. ADR-0014 later changed `auto` to select `.NET` only.
- During early migration, `auto` could prefer PowerShell to preserve current behavior. After the major command set became stable in .NET, ADR-0014 removed PowerShell fallback.

The provider contract includes:

- `Name()`
- `Supports(command)`
- `Info(ctx)`
- `Execute(ctx, request)`

The bridge protocol is stdin/stdout JSON. The child bridge process receives one request JSON document on stdin and writes one response JSON document to stdout. stdout is reserved for protocol JSON only; logs must be carried in response fields or stderr. This rule applies to both normal results and fatal bridge failures.

Every bridge request and response includes `protocol_version` and `request_id`. The Go side must reject incompatible protocol versions before trusting command output.

## Consequences

- Positive: `.NET` migration can happen command by command without breaking unsupported commands in `auto`.
- Positive: Explicit `--bridge dotnet` is reliable for proving that PowerShell was not used.
- Positive: Protocol version checks prevent accidental mixing of incompatible `xlflow.exe` and `xlflow-excel-bridge.exe`.
- Positive: stdout JSON-only behavior keeps machine output stable for agents and scripts.
- Negative: Go command plumbing must carry bridge mode through many workbook-backed commands.
- Negative: While both bridges exist, tests must cover provider selection, unsupported-command failures, and fallback paths.
- Negative: Default `auto` behavior changes require careful release notes because users may observe a different selected bridge after .NET becomes preferred.

## Alternatives Considered

1. **Always fallback to PowerShell, even for `--bridge dotnet`** — Rejected because it hides whether the .NET bridge actually works in locked-down environments.
2. **Switch all commands to .NET at once** — Rejected because pull/push/run/test/form/export-image have different risk profiles and need staged Excel COM E2E.
3. **Use environment variable only, without CLI/config support** — Rejected because users need explicit per-command control and project-level defaults.
4. **Let bridges emit native xlflow envelopes without a provider contract** — Rejected because protocol compatibility and fallback decisions would become implicit and harder to test.

## Related

- `docs/adr/ADR-0008-dotnet-excel-bridge.md`
- `docs/bridge/bridge-protocol.md`
- `docs/bridge/dotnet-bridge.md`
- `docs/specs/cli-contract.md`
- `docs/design.md`
