# ADR-0008: .NET Excel Bridge for Windows Automation

## Status

Accepted

## Context

xlflow's first Excel automation boundary is the bundled PowerShell bridge described in ADR-0001. That approach kept the Go CLI small and made early automation scripts easy to inspect, but it has limits that matter more as xlflow targets corporate and agent-driven workflows:

- Corporate PowerShell execution policies, endpoint controls, or missing optional PowerShell features can block workbook-backed commands.
- Excel COM, VBIDE, Win32 dialog watching, process ownership, and clipboard handling are easier to express and test in C# than in large PowerShell scripts.
- Prior runtime diagnostics, modal dialog handling, session reuse, and UserForm work have made the PowerShell bridge increasingly large and harder to keep safe.
- The Go CLI should stay cross-platform for source-only commands, while Windows workbook automation remains explicitly Windows/Excel-specific.

We need a durable migration path that improves Windows automation without forcing a big-bang rewrite or removing the existing bridge before the replacement is proven.

## Decision

Introduce a separate C#/.NET executable named `xlflow-excel-bridge.exe` as the planned Windows Excel automation bridge.

Responsibility split:

- Go core owns CLI command routing, config loading, source tree management, lint/format/static analysis, output envelope mapping, release packaging, and bridge provider selection.
- The .NET bridge owns Excel COM automation, VBIDE automation, workbook/session handling, macro execution, UserForm import/export, Win32 dialog watching, runtime/compile error capture, image export fallback, process/window control, and clipboard control.
- The existing PowerShell bridge remains a legacy fallback while the .NET bridge reaches parity. ADR-0014 later removed this fallback after parity was reached.

The .NET bridge runs as a separate process beside `xlflow.exe`. The Go CLI invokes it through the bridge provider contract described in ADR-0009. Windows release archives will eventually include both `xlflow.exe` and `xlflow-excel-bridge.exe`; non-Windows archives keep the Go CLI only.

The .NET bridge must treat Excel COM as STA-bound. COM automation must run on an STA entrypoint or dedicated STA dispatcher. Handler code must not `await` and then resume Excel COM calls on a ThreadPool/MTA thread. Initial COM-facing command handlers should be synchronous unless a dedicated STA dispatcher is introduced.

PowerShell was not removed by this decision. ADR-0014 later removed PowerShell bridge selection and fallback after the .NET bridge reached parity.

## Consequences

- Positive: Corporate environments that block PowerShell scripts have a path to workbook-backed commands through a normal executable.
- Positive: Win32, COM, UI Automation, process, and clipboard handling can move into a language/runtime better suited to those APIs.
- Positive: The Go CLI remains small and source-only commands remain cross-platform.
- Positive: The migration can proceed command by command because PowerShell remains as fallback.
- Negative: Windows releases gain another executable and another build/publish path.
- Negative: Some corporate environments can still block unsigned or unknown executables through AppLocker, WDAC, Defender, EDR, or code-signing policy.
- Negative: COM apartment threading becomes a hard architectural constraint for C# implementation and tests.
- Negative: During migration, both PowerShell and .NET bridge behavior must be kept compatible enough for users and CI to reason about fallback.

## Alternatives Considered

1. **Keep PowerShell as the only bridge** — Rejected because execution policy and long-term maintainability risks remain.
2. **Rewrite Excel automation directly in Go** — Rejected because Go has weaker ergonomics for COM/VBIDE/Win32/UI Automation work and would likely recreate lower-level interop complexity in the CLI process.
3. **Embed .NET in `xlflow.exe`** — Rejected because it couples the Go CLI to a Windows runtime and complicates non-Windows distribution.
4. **Start with a daemon or gRPC bridge** — Rejected for the initial migration because command-scoped stdin/stdout JSON is simpler to debug, package, and fallback.

## Related

- `docs/adr/ADR-0001-agent-ready-vba-cli-architecture.md`
- `docs/adr/ADR-0004-explicit-excel-session-mode.md`
- `docs/adr/ADR-0009-bridge-provider-contract.md`
- `docs/bridge/bridge-protocol.md`
- `docs/bridge/dotnet-bridge.md`
- `docs/design.md`
