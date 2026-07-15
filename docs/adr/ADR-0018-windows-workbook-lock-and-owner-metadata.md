# ADR-0018: Windows Workbook Lock and Observational Owner Metadata

## Status

Accepted

## Context

ADR-0016 defines the authoritative command-policy registry and canonical
workbook identity, but intentionally defers the operating-system primitive used
to coordinate independent xlflow processes. Workbook operations can originate
from direct Windows commands, WSL-delegated commands, editor integrations, and
automation. An in-process mutex cannot coordinate those callers, and the
existence of a lock or metadata file cannot safely represent ownership after a
crash.

Contention diagnostics also benefit from owner information such as the process,
command, operation kind, and start time. That information can be missing,
malformed, or left behind after abnormal termination, so it cannot participate
in the ownership decision. Publishing and reading it without synchronization
would additionally allow a new owner to be misidentified using the previous
owner's metadata.

## Decision

On Windows, xlflow coordinates each non-parallel-safe workbook operation with a
non-blocking byte-range lock acquired through `LockFileEx`. The lock file is
derived from the opaque `WorkbookIdentity.LockID` and stored in the current
user's xlflow coordination directory. The path never contains the workbook
path. The held operating-system lock is the sole authority for ownership; file
existence and metadata contents are never authoritative.

The operation lock is held for the complete CLI operation, including backups,
bridge invocation, Excel and VBIDE work, saves, rollback, and cleanup. Windows
releases the lock when its owning process terminates. The lock layer also
supports context-aware polling so later explicit wait modes can add cancellation
without changing the ownership primitive. Immediate failure remains the default
policy in this decision.

Source-only commands and workbook observers classified as parallel-safe do not
acquire the operation lock. A non-parallel-safe workbook command must resolve a
canonical workbook identity before its body starts. WSL workbook commands remain
covered by ADR-0011: the delegated Windows process acquires and owns the lock.
`excel_instance` coordination is not implemented by treating it as a workbook
lock.

Each successful acquisition publishes observational owner metadata in the same
per-user coordination directory. The metadata uses a versioned schema and
contains an acquisition generation token, canonical workbook path, PID, stable
command identifier, operation kind, resource scope, and UTC start time. It does
not persist raw command arguments. Metadata is written by atomic replacement and
normal cleanup removes it only when its generation token still matches the
releasing lease. Publication failure aborts acquisition before workbook work can
start; cleanup failure does not turn an otherwise successful workbook operation
into a failure.

The lock file reserves distinct byte ranges for the authoritative operation lock
and a metadata publication guard. A writer holds the publication guard while it
acquires the operation lock and publishes metadata. A contending reader first
observes operation-lock contention, then holds the publication guard, confirms
that contention still exists, and only then reads metadata. Release performs
token-conditional cleanup under the publication guard before releasing the
operation lock. This handshake prevents both stale crash metadata and the brief
new-owner publication window from being reported as current ownership.

Malformed, missing, partial, or stale metadata is ignored. A caller still
receives `workbook_busy` based solely on operation-lock contention, with owner
details omitted when no trustworthy current record is available.

The detailed acquisition, metadata, and diagnostic contracts live in
`docs/specs/workbook-coordination.md` and `docs/specs/cli-contract.md`.

## Consequences

- Positive: independent Windows and WSL-delegated processes cannot overlap
  conflicting operations on the same canonical workbook identity.
- Positive: process termination releases authoritative ownership without stale
  lock cleanup.
- Positive: different workbook identities remain independently executable.
- Positive: owner details improve diagnostics without weakening crash safety or
  becoming an ownership oracle.
- Positive: later cancellation-aware wait options can reuse the lock layer.
- Negative: coordination depends on Windows byte-range locking semantics and a
  writable per-user local state directory.
- Negative: polling-based waiting is less direct than a waitable named kernel
  object and must balance cancellation responsiveness with wake-up overhead.
- Negative: metadata publication adds another byte-range lock and filesystem
  operations to acquisition and release.
- Negative: commands classified at `excel_instance` scope still need a separate
  coordination decision.
- Deferred: public wait flags, timeout behavior, FIFO fairness, and coordination
  fields in `session status` remain follow-up work.

## Alternatives Considered

1. **Windows named mutex** - Rejected because cancellation-aware acquisition and
   ownership diagnostics would require a separate metadata protocol while Go
   callers also need careful thread-affinity handling around mutex ownership.
2. **Lock-file existence** - Rejected because an abnormal exit can leave the
   file behind and permanently or incorrectly block future work.
3. **Metadata as ownership** - Rejected because PID reuse, partial writes, and
   crash leftovers cannot provide authoritative ownership.
4. **Lock only around bridge calls** - Rejected because backups, rollback, pack,
   saves, and cleanup can access the same workbook outside a bridge invocation.
5. **Store coordination state in each project's `.xlflow` directory** - Rejected
   because two projects or working directories can address the same workbook and
   must converge on one lock and metadata location.
6. **Read metadata immediately after contention without a guard** - Rejected
   because a new owner may hold the operation lock before replacing a previous
   owner's metadata.

## Related

- `docs/adr/ADR-0016-workbook-operation-coordination.md`
- `docs/adr/ADR-0011-wsl-windows-cli-delegation.md`
- `docs/specs/workbook-coordination.md`
- `docs/specs/cli-contract.md`
- `internal/coordination`
- xlflow issues #311, #320, #321, and #323
