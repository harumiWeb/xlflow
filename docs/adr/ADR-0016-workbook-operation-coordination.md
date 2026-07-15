# ADR-0016: Workbook Operation Coordination Boundaries

## Status

Accepted

## Context

xlflow can reach the same workbook through independent CLI processes, a reused
Excel session, the .NET bridge, VS Code commands, AI agents, automated scripts,
and a Windows `xlflow.exe` delegated from WSL. Excel COM, VBIDE, VBA execution,
and the UserForm Designer are not general-purpose concurrent environments.
Allowing two non-parallel-safe operations to overlap can interleave mutations,
expose partially updated state, or produce non-deterministic automation errors.

Command implementations currently imply their own workbook and concurrency
requirements. Adding coordination through scattered command lists would let a
new command bypass protection and would allow the CLI, bridge, and integrations
to disagree about command safety. Session IDs are also an insufficient resource
boundary: xlflow can automatically reuse a matching session, while another
process can independently address the same workbook without naming that
session.

Coordination also needs a stable key. Relative and absolute paths, Windows path
casing and separators, WSL path translation, and equivalent lexical path forms
can all identify the same workbook. Raw paths are unsuitable as operating-system
synchronization names because they can expose project locations and contain
characters unsupported by the selected primitive.

This ADR decides the policy and identity boundaries. The operating-system lock
primitive, busy diagnostics, waiting behavior, and status reporting are deferred
to later work.

## Decision

The Go core owns one authoritative command coordination registry. Each executable
command has an explicit policy describing its resource scope, operation kind,
parallel safety, busy retryability, and default wait policy. The registry can be
queried by stable command ID, Cobra leaf path, and .NET bridge command selector,
including action fields needed to distinguish multiplexed bridge requests.

An executable command without an explicit policy fails closed with a typed
`coordination_policy_missing` error. It does not inherit a permissive default.
Policy consumers use the Go registry; the .NET bridge does not maintain a second
command list. Future external discovery may serialize this registry, but this
decision does not add or change a public CLI, bridge, or capabilities schema.

Coordination is keyed by canonical workbook identity rather than session ID. An
identity contains:

- a canonical Windows path retained for human-readable diagnostics; and
- an opaque, domain-separated SHA-256 lock identifier that does not contain the
  workbook path.

Identity construction takes both the base directory and workbook path explicitly
so it does not depend on the process working directory. It normalizes the path
lexically, attempts to resolve existing symbolic-link or junction aliases, and
uses Windows case-insensitive comparison semantics while preserving UNC path
semantics. Failure to resolve a link does not require the workbook to exist and
falls back to the normalized lexical path.

ADR-0011 remains the WSL boundary. WSL translates supported `/mnt/<drive>/...`
paths with `wslpath -w` and delegates the complete command to Windows
`xlflow.exe`. The Windows process creates the identity from the translated
Windows path. The canonicalizer does not independently interpret WSL paths.

The exact cross-process operating-system lock primitive is intentionally not
selected here. A later decision or implementation must use a real process-safe
primitive as the ownership authority; diagnostic files or session metadata must
never become the source of truth for ownership.

The detailed policy and identity contracts live in
`docs/specs/workbook-coordination.md`.

## Consequences

- Positive: CLI dispatch, bridge invocation, and future integrations can reason
  from one command safety source of truth.
- Positive: fail-closed coverage prevents new executable commands from silently
  bypassing future coordination.
- Positive: independent processes and auto-reused sessions converge on the same
  workbook resource boundary.
- Positive: opaque lock identifiers are deterministic and do not disclose full
  workbook paths in synchronization primitive names.
- Positive: identity generation does not require Excel, an open workbook, or an
  existing workbook file.
- Negative: every executable command and multiplexed bridge action must be kept
  registered as the command surface evolves.
- Negative: conservative policies can serialize operations that might eventually
  prove safe to run together.
- Negative: lexical normalization cannot reliably unify mapped drives with UNC
  paths, DFS or DNS aliases, hard links, 8.3 names, or every network alias.
- Negative: the initial case-insensitive identity contract does not distinguish
  Windows directories explicitly configured for case-sensitive lookup.
- Deferred: cross-process acquisition, crash recovery, busy/timeout diagnostics,
  opt-in waiting, and session status integration require subsequent work.

## Alternatives Considered

1. **Keep policy lists in each CLI, bridge, and editor integration** - Rejected
   because the lists would drift and new commands could omit safety handling.
2. **Use a permissive policy for unregistered commands** - Rejected because an
   omission would silently weaken coordination exactly when a command is new.
3. **Coordinate by session ID** - Rejected because callers can address or
   auto-reuse the same workbook without sharing an explicit session ID.
4. **Use the canonical path directly as the lock name** - Rejected because paths
   disclose project locations and may violate synchronization-name constraints.
5. **Resolve identity from workbook file handles or file IDs only** - Rejected
   because coordination must work before the workbook exists or is open, and
   network filesystems do not provide a uniform stable identity.
6. **Convert WSL paths inside the canonicalizer** - Rejected because ADR-0011
   already assigns path translation and Excel command ownership to the delegated
   Windows CLI.
7. **Choose a Windows named mutex or file-lock implementation now** - Deferred
   until the lock sub-issue can compare lifecycle, cancellation, diagnostics,
   crash recovery, and WSL-delegated process behavior together.

## Related

- `docs/specs/workbook-coordination.md`
- `docs/specs/cli-contract.md`
- `internal/coordination`
- `docs/adr/ADR-0004-explicit-excel-session-mode.md`
- `docs/adr/ADR-0008-dotnet-excel-bridge.md`
- `docs/adr/ADR-0009-bridge-provider-contract.md`
- `docs/adr/ADR-0011-wsl-windows-cli-delegation.md`
- xlflow issues #311, #318, and #319
