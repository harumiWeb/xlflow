# ADR-0011: WSL Development Through Windows xlflow Delegation

## Status

Accepted

## Context

AI agents and developers often keep Git, editors, and Linux tooling in WSL, but
Excel, VBA, COM, VBIDE, and UI Automation remain Windows-only dependencies.
Running the Windows `.NET` bridge directly from the WSL Go process would split
CLI responsibility across operating systems and duplicate path, output, and
exit-code handling already owned by the Windows CLI.

Projects stored only in the WSL filesystem are not reliably visible to Excel.
Windows-mounted paths such as `/mnt/c/...` provide a stable shared filesystem
boundary.

## Decision

Treat WSL as a development frontend and delegate Excel-related commands to an
installed Windows `xlflow.exe`.

- Detect WSL from `WSL_INTEROP`, `WSL_DISTRO_NAME`, or `/proc/version`.
- Support delegated projects only under `/mnt/<drive>/...`.
- Resolve Windows xlflow from `XLFLOW_WINDOWS_EXE`, then `xlflow.exe` on the
  interoperable PATH.
- Translate WSL absolute path arguments with `wslpath -w`; keep relative paths
  relative to the shared project working directory.
- Delegate the complete Windows CLI command rather than invoking the `.NET`
  bridge directly. Windows xlflow continues to own configuration, preflight,
  bridge selection, output envelopes, and Excel automation.
- Preserve stdin, stdout, stderr, and the Windows process exit code. `doctor` is
  the exception: WSL captures Windows JSON, adds host, executable, bridge, Excel,
  and path-translation diagnostics, then renders the combined envelope.
- Add xlflow runtime variables such as `XLFLOW_MODE` and
  `XLFLOW_EXCEL_BRIDGE` to `WSLENV` for the delegated Win32 process.
- Keep source-only commands local to WSL. Help and shell completion are never
  delegated.
- Set an internal environment marker on child processes to prevent recursive
  delegation.

The Windows and WSL xlflow versions may differ. `doctor` reports the mismatch as
a warning but does not block delegation.

## Consequences

- Positive: Agents can perform edit, push, run, test, inspect, and session
  workflows from WSL while using real Windows Excel.
- Positive: JSON and exit-code behavior remains defined by the existing Windows
  CLI contract.
- Positive: The `.NET` bridge remains a Windows implementation detail and does
  not gain a second WSL-specific protocol.
- Negative: Users must install xlflow on both WSL and Windows.
- Negative: Projects under `/home`, WSL virtual disks, or other
  Windows-invisible paths are rejected for delegated commands.
- Negative: Absolute path-bearing CLI arguments require explicit translation
  coverage as new path options are added.
- Negative: End-to-end verification requires a real WSL, Windows, and Excel
  environment in addition to cross-platform unit tests.

## Alternatives Considered

1. **Run the Windows `.NET` bridge directly from WSL** — Rejected because it
   bypasses Windows CLI preflight, provider selection, output mapping, and
   compatibility checks.
2. **Delegate every command to Windows** — Rejected because source linting,
   formatting, analysis, Git, and agent tooling should retain native WSL paths
   and performance.
3. **Copy WSL-only projects to a temporary Windows directory** — Rejected
   because synchronization, file identity, generated state, and failure
   recovery would become implicit and unsafe.
4. **Add a daemon or network service on Windows** — Rejected because WSL
   interoperability can launch the existing executable without introducing a
   persistent service or authentication surface.

## Related

- `docs/adr/ADR-0001-agent-ready-vba-cli-architecture.md`
- `docs/adr/ADR-0008-dotnet-excel-bridge.md`
- `docs/adr/ADR-0009-bridge-provider-contract.md`
- `docs/specs/cli-contract.md`
- `vitepress/installation.md`
