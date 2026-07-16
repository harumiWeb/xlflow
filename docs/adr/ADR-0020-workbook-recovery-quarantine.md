# ADR-0020: Workbook Recovery Quarantine After Indeterminate Excel Termination

## Status

Accepted

## Context

ADR-0018 makes the Windows operating-system workbook lock the sole authority
for an actively executing xlflow command. That lock is released automatically
when its owning process exits, which is the correct crash-safe ownership
behavior under normal completion.

Some Excel automation failures have a different lifetime. A macro timeout,
terminated bridge worker, fatal COM/RPC disconnect, incomplete cleanup, or
poisoned session can cause the xlflow command to return while Excel or VBA may
still be running or mutating the workbook. Releasing the process-owned lock in
that state allows another xlflow process to acquire it even though the
Excel-side operation may not have finished.

A stale lock file cannot safely extend ownership beyond its process, and owner
metadata cannot become an ownership oracle without violating ADR-0018. xlflow
therefore needs a separate, machine-readable safety state that prevents new
workbook work without pretending that a process still owns the lock.

## Decision

xlflow persists a versioned recovery-required marker for a canonical workbook
identity when an explicitly classified outcome means Excel-side execution or
cleanup cannot be confirmed complete. The public blocked-operation code is
`workbook_recovery_required`.

The marker is diagnostic authorization state, not active ownership. The
operating-system workbook lock remains the only ownership authority. Every
workbook command that is subject to recovery policy acquires the normal lock
first, then checks the marker before entering its command body. An uncertain
operation publishes the marker before releasing its lock, and a recovery
operation clears it only while holding that lock.

The Go coordination layer owns marker storage and the central command policy.
Every executable descriptor declares one of four recovery behaviors:
`not_applicable`, `block`, `observe`, or `recover`. Workbook reads, mutations,
execution, and Designer automation normally block. Status and process-list
operations may observe without accessing unsafe live workbook state. Explicit
recovery operations include managed-session discard, successful process
cleanup, and `recovery clear`.

Recovery markers use the same canonical workbook identity and per-user Windows
coordination directory as the operation lock, but a separate atomic JSON file.
Malformed or unsupported recovery metadata fails closed. Marker files never
participate in lock ownership, busy retryability, or stale-owner inference.

`--wait` applies only to current operating-system lock contention. It does not
wait for, bypass, or clear recovery-required state. Once the command acquires
the lock and observes quarantine, it returns immediately with
`workbook_recovery_required`.

Normal recovery clearing requires positive evidence that the affected Excel
process is no longer running, or successful completion of a recovery operation
that closes or terminates it. `recovery clear --force` may remove the marker
without verification only through an explicit user action and must warn that it
does not terminate VBA, close Excel, repair the workbook, or prove safety.

Managed sessions in recovery-required state cannot be saved. A successful
`session stop --discard` may clear the marker after xlflow confirms the managed
Excel process ended. Externally owned Excel sessions are detached without
claiming that xlflow closed or discarded the user's workbook.

WSL continues to delegate workbook commands to Windows under ADR-0011. The
delegated Windows process observes and updates the same Windows-side recovery
state as direct Windows invocations.

The detailed metadata, publication, diagnostics, command classification,
session, process cleanup, and status contracts live in
`docs/specs/workbook-coordination.md`, `docs/specs/cli-contract.md`, and
`docs/bridge/dotnet-bridge.md`.

## Consequences

- Positive: a released process lock no longer authorizes unsafe follow-up work
  when Excel or VBA may still be active.
- Positive: busy ownership and recovery-required safety state remain distinct
  and independently observable.
- Positive: CLI, direct Runner, WSL delegation, editor integrations, scripts,
  and AI agents receive one stable recovery contract.
- Positive: recovery publication and clearing are serialized with normal
  workbook operations.
- Negative: a marker can require manual recovery even when Excel eventually
  became safe after xlflow lost the ability to verify it.
- Negative: malformed metadata and storage-read failures block unsafe commands
  until the state can be inspected or force-cleared.
- Negative: verified recovery depends on available Excel process identity; when
  no PID was captured, cleanup-all or explicit force clear may be necessary.
- Negative: older xlflow versions do not understand recovery markers and cannot
  participate safely in mixed-version execution.
- Limitation: a process crash before it publishes the recovery marker remains
  undetectable; stale process metadata is not promoted into authority to close
  that gap.
- Limitation: recovery state is per Windows user and machine, not distributed
  across users, machines, or general `excel_instance` coordination.

## Alternatives Considered

1. **Keep the operating-system lock after the command returns** - Rejected
   because the lock is owned by the process and must be crash-released; keeping
   it would require a persistent owner process outside the current design.
2. **Treat a stale lock or owner-metadata file as ownership** - Rejected because
   file existence, PID reuse, partial writes, and crash leftovers cannot prove
   current ownership and would contradict ADR-0018.
3. **Automatically clear recovery when the recording process exits** - Rejected
   because that process may be only the bridge worker while Excel and VBA remain
   alive.
4. **Make `--wait` poll recovery state** - Rejected because recovery requires an
   explicit safety transition, not ordinary lock-contention retry.
5. **Allow all read-only workbook commands during recovery** - Rejected because
   many apparent reads still attach to the uncertain Excel instance, invoke COM,
   or observe a workbook that may be changing asynchronously.
6. **Automatically terminate every affected Excel instance** - Rejected because
   external Excel instances may be user-owned and contain unrelated unsaved
   workbooks.

## Related

- `docs/adr/ADR-0011-wsl-windows-cli-delegation.md`
- `docs/adr/ADR-0016-workbook-operation-coordination.md`
- `docs/adr/ADR-0018-windows-workbook-lock-and-owner-metadata.md`
- `docs/adr/ADR-0019-opt-in-workbook-wait-policy.md`
- `docs/specs/workbook-coordination.md`
- `docs/specs/cli-contract.md`
- `docs/specs/runtime-debugging.md`
- `docs/bridge/dotnet-bridge.md`
- `internal/coordination`
- xlflow issues #311 and #335
