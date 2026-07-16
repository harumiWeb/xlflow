# Workbook Operation Coordination

This spec defines the command-policy, canonical workbook-identity,
cross-process lock, owner-metadata, and recovery-quarantine contracts used by
workbook operation coordination. Public waiting options and coordination status
output build on these contracts.

See ADR-0016 for the policy and identity rationale and ADR-0018 for the Windows
lock and metadata decisions. ADR-0020 defines recovery quarantine after
indeterminate Excel termination.

## Command Policy Contract

The Go core owns one authoritative registry of executable command descriptors.
Each descriptor has a stable command ID, selectors used by the CLI and bridge,
and one coordination policy. Consumers query this registry instead of maintaining
hardcoded workbook-command lists.

The policy vocabulary is:

| Field                 | Values                                          | Meaning                                                                        |
| --------------------- | ----------------------------------------------- | ------------------------------------------------------------------------------ |
| `resource_scope`      | `none`, `workbook`, `excel_instance`            | Smallest resource boundary required by the operation.                          |
| `operation_kind`      | `read`, `mutate`, `execute`, `designer`         | The operation's highest-risk behavior.                                         |
| `parallel_safe`       | boolean                                         | Whether the operation may proceed without exclusive coordination at its scope. |
| `retryable_when_busy` | boolean                                         | Whether a future caller may opt into retrying after resource contention.       |
| `default_wait_policy` | `fail`, `wait`                                  | Whether acquisition should fail immediately or wait by default.                |
| `recovery_behavior`   | `not_applicable`, `block`, `observe`, `recover` | How the command behaves when workbook recovery is required.                    |

All policies use `default_wait_policy: fail`. Callers may opt into bounded
waiting through the public CLI contract below. A resource operation is retryable only when it is non-parallel-safe
and retrying after contention is meaningful. `parallel_safe: true` means that the
operation does not need exclusive acquisition at its declared scope; it does not
introduce a shared read-lock mode.

Every descriptor must also declare `recovery_behavior`; there is no permissive
zero value. `not_applicable` is limited to operations that do not address unsafe
workbook or Excel state. `block` requires a lease-scoped marker check before the
command body starts. `observe` permits metadata-only status or process
observation. `recover` permits an explicit recovery operation while preserving
normal lock serialization.

When flags can change a command's risk, the descriptor uses the most restrictive
behavior supported by that executable command unless the registry has an explicit
selector that distinguishes the variants. WSL delegation classification is
separate from coordination policy: it determines where a command runs, not what
resource safety the command requires.

### Command Selectors

Every executable Cobra leaf command resolves to exactly one descriptor by its
normalized leaf path. Command containers, help, and generated completion commands
are not executable policy entries.

Bridge-backed commands resolve through the same registry before invocation. A
selector contains the bridge command and, for multiplexed bridge endpoints, the
request field such as `Action` that identifies the actual operation. The .NET
bridge must not duplicate the registry. Stable command IDs are internal identifiers
in this version; they are suitable for future serialization but are not yet a
public capabilities schema.

Registry enumeration is deterministic and returns defensive copies so callers
cannot mutate the authoritative definitions.

### Fail-Closed Behavior

An executable CLI command or bridge invocation without an explicit matching
descriptor returns a typed policy lookup failure. The CLI error code is
`coordination_policy_missing`. There is no implicit `resource_scope: none`,
`recovery_behavior: not_applicable`, or other permissive fallback.

Coverage tests enumerate executable Cobra leaves and used bridge command/action
selectors. Adding a command without adding its policy must fail those tests.

### Initial Classification

The registry classifies commands conservatively:

- Source-only operations such as lint, format, analyze, LSP, source inspection,
  `test list`, and `form new` use `resource_scope: none` and
  `recovery_behavior: not_applicable`.
- Workbook inspection and synchronization reads, including pull, init, macro and
  form listing, formula pull, workbook diff, and workbook/sheet/range inspection, use
  `resource_scope: workbook`, `operation_kind: read`, are not parallel-safe, and
  use `recovery_behavior: block`.
- Workbook creation and mutation, including new, push, rollback, save, session
  start/stop, runner changes, pack artifact generation, and worksheet/cell/UI edits, use
  `resource_scope: workbook`, `operation_kind: mutate`, and are not parallel-safe.
  Normal mutation commands use `recovery_behavior: block`; `session stop` is
  recovery-capable so it can discard an unsafe managed session.
- VBA run and test use `resource_scope: workbook`, `operation_kind: execute`, and
  are not parallel-safe. Executing variants use `recovery_behavior: block`;
  source-only `test list` remains not applicable.
- UserForm migration, snapshot, build, specification application, image export,
  and Designer inspection use `resource_scope: workbook`,
  `operation_kind: designer`, are not parallel-safe, and use
  `recovery_behavior: block`.
- Top-level status and session status are read-only observers. They use
  `resource_scope: workbook`, `operation_kind: read`, and `parallel_safe: true`
  with `recovery_behavior: observe` so they can report busy and recovery state
  without first waiting for either state to clear.
- Environment checks, active-workbook attachment, and Excel process cleanup use
  `resource_scope: excel_instance` when they inspect or mutate state broader than
  one configured workbook. Process list observes recovery metadata, while
  process cleanup is recovery-capable. Attachment remains blocked when it would
  access a quarantined workbook.
- `recovery clear` uses workbook scope and `recovery_behavior: recover`.

### UserForm and Designer Coverage

The configured workbook lock is the only exclusion primitive for UserForm work;
there is no separate Designer lock. The Designer-classified command IDs are
`form.migrate.sidecar`, `form.snapshot`, `form.build`, `form.apply`,
`form.export-image`, and `inspect.form`. `list.forms` and `pull` use workbook
read policy, while `push` uses workbook mutate policy. All nine operations use
the same canonical workbook identity and therefore conflict with Designer,
execution, test, synchronization, and other workbook operations on that file.

The CLI lease surrounds the complete leaf handler. For migration and snapshot
this includes Designer inspection plus sidecar/spec writes; for build and push
it includes VBIDE mutation, `.frm`/`.frx` synchronization, save/restore, and
cleanup. Direct Runner calls for `form-write`, `inspect-form`,
`form-export-image`, list/forms, pull, and push resolve the same central
descriptors before starting a bridge provider. CLI-owned Runner calls receive a
borrowed `LeaseSet` from the outer command rather than bypassing coordination.
Direct Runner calls acquire their own leases. Both paths use the same
lease-scoped recovery check and publication API.

`form.new`, source rename, and source delete remain `resource_scope: none` and
do not open or lock a workbook. Their later workbook application occurs through
coordinated `push`. Under WSL the source-only steps remain local; workbook-backed
UserForm commands are delegated and acquire the lock in the Windows CLI.

## Canonical Workbook Identity

A workbook identity contains:

- `CanonicalPath`, a human-readable canonical Windows path for diagnostics; and
- `LockID`, an opaque deterministic identifier safe for a future cross-process
  synchronization primitive.

The API requires an explicit base directory and workbook path. A relative workbook
path is resolved against that base; identity generation never reads the process
working directory implicitly and does not require the workbook to be open or to
exist.

### Normalization

Identity generation applies these steps in order:

1. Resolve a relative workbook path against the explicit base directory and make
   the result absolute.
2. Clean redundant separators and `.` or `..` path segments.
3. Starting with the full path, walk toward the root until the nearest existing
   ancestor can be resolved. Resolve symbolic links and junctions in that
   ancestor, then append the not-yet-created tail in its original order. If no
   ancestor can be resolved, retain the lexical absolute path.
4. Convert Windows extended path prefixes to their normal drive or UNC form while
   preserving UNC semantics.
5. Normalize separators to the canonical Windows form.
6. Normalize drive-letter casing for the diagnostic path.
7. Build a case-insensitive comparison key from the canonical path.

Equivalent path casing, slash direction, drive-letter casing, relative/absolute
forms, and redundant lexical segments therefore produce the same comparison key.
UNC paths retain their leading UNC semantics and are not converted to local-drive
paths.

`CanonicalPath` preserves a useful diagnostic representation; callers must use
`LockID`, not display-path string equality, as the synchronization key.

### Lock Identifier

`LockID` is the lowercase hexadecimal SHA-256 digest of a domain separator plus
the normalized comparison key, with a fixed ASCII prefix identifying xlflow
workbook coordination. The domain separator prevents accidental reuse of the same
digest namespace for another resource type. The resulting identifier contains
only fixed ASCII prefix characters and hexadecimal digits and never includes the
workbook path.

SHA-256 collision risk is negligible for practical workbook coordination. The
operating-system lock implementation may add a primitive-specific filename
suffix but must preserve the opaque identity and must not embed `CanonicalPath`
in the primitive name.

### WSL Boundary

The canonicalizer accepts Windows paths; it does not translate `/mnt/...` or
other Linux paths. Under ADR-0011, WSL translates supported absolute paths with
`wslpath -w` and delegates the complete command to Windows `xlflow.exe`. Identity
generation then runs in the Windows process using the translated path. A supported
WSL-mounted path and its direct Windows representation therefore share an identity.

WSL-only paths such as `/home/...` remain unsupported for delegated workbook
commands and fail in the delegation layer before identity generation.

### Unresolved Aliases

Best-effort link resolution does not promise that every path alias converges.
Unless Windows resolves them through the normal path/link resolution above, the
following may produce distinct identities even when they ultimately reach the
same file:

- a mapped drive and its backing UNC share;
- DFS targets, DNS aliases, and alternate server names;
- hard links and 8.3 short names;
- network redirector aliases;
- paths that differ only by Unicode normalization; and
- paths in directories explicitly configured for case-sensitive lookup.

An existing symlink or junction parent is resolved even when the workbook or
intermediate children beneath it do not exist yet. A broken, inaccessible, or
not-yet-created alias itself uses its lexically normalized absolute form.
Diagnostic output may expose the canonical workbook path, but synchronization
primitive names must expose only `LockID`.

## Cross-Process Workbook Lock

### Applicability and Lifetime

A command acquires the workbook operation lock when its resolved descriptor has
`resource_scope: workbook` and `parallel_safe: false`. Source-only operations
with `resource_scope: none` and parallel-safe workbook observers do not acquire
it. `resource_scope: excel_instance` requires a separate coordination primitive
and must not silently fall back to a workbook lock.

An applicable command must resolve its `WorkbookIdentity` before entering the
command body. Acquisition precedes all workbook-bound work, and the lease is
held until backups, bridge calls, Excel and VBIDE activity, saves, rollback, and
cleanup have finished. Validation that does not access workbook state may happen
before acquisition.

Immediate acquisition is the default. If the authoritative byte range is
already locked, acquisition returns a typed busy result without entering the
command body. The internal API also polls with a context for the explicit,
cancellation-aware wait contract below.

### Windows Primitive and Storage

Windows uses `LockFileEx` and `UnlockFileEx` byte-range locks on a lock file
derived from `WorkbookIdentity.LockID`. Coordination files live in a per-user
xlflow state directory under the Windows local application-data location, not in
the project `.xlflow` directory. This lets independent projects and working
directories that address the same workbook converge on one resource boundary.

The lock file reserves separate fixed byte ranges for:

- the authoritative workbook operation lock; and
- the metadata publication guard.

The presence of the lock file, its timestamps, and its contents are not evidence
of ownership. Only successful operating-system byte-range acquisition establishes
ownership. Closing the lease handle or terminating the owner process releases
the authoritative lock. Lock and metadata filenames expose `LockID` only and
must not contain the canonical workbook path.

WSL does not acquire a Linux-side lock. Under ADR-0011, a workbook command is
delegated in full and the Windows `xlflow.exe` acquires the same lock that a
direct Windows invocation would use.

### Acquisition Outcomes

Acquisition distinguishes:

- acquired: the caller owns a lease and may enter the command body;
- busy: another process owns the operation byte range; and
- operational failure: coordination storage or the Windows locking API failed.

Busy is retryable only when the descriptor's `retryable_when_busy` field is true.
It is not inferred from metadata. Different `LockID` values use different lock
files and do not contend.

## Owner Metadata

Owner metadata is observational data associated with a currently held operation
lock. Its versioned schema contains:

```json
{
  "schema_version": 1,
  "generation": "acquisition-specific-random-token",
  "workbook": "C:\\projects\\sample\\sample.xlsm",
  "pid": 12345,
  "command": "run",
  "operation_kind": "execute",
  "resource_scope": "workbook",
  "started_at": "2026-07-15T09:30:00Z"
}
```

`command` is a stable command ID or normalized leaf path. Raw argv and command
arguments are not persisted because they may contain secrets or sensitive file
locations. `workbook` is the canonical diagnostic path and `started_at` is UTC.
`generation` is newly generated for each successful acquisition.

Metadata is written to a temporary file and atomically replaces the published
file. Normal release deletes the published file only when its generation still
matches the releasing lease. A missing, malformed, unsupported, partial, or
generation-mismatched record is ignored. Publication failure aborts acquisition
before the command body starts so stale metadata cannot be paired with a new
owner. Cleanup failure after the command body is best-effort and must not fail
an otherwise successful workbook operation. Neither failure mode makes metadata
authoritative.

### Publication Handshake

Writers and readers use the metadata publication guard to avoid misidentifying
an owner across handoff:

1. An acquiring writer locks the publication guard, acquires the operation byte
   range, publishes its metadata, and releases the guard.
2. A contending reader first observes operation-lock contention, then acquires
   the publication guard and confirms that the operation remains busy.
3. Only after that confirmation may the reader parse the published metadata.
4. A releasing owner acquires the publication guard, conditionally removes its
   matching generation, releases the operation lock, and releases the guard.

Metadata may remain after a crash, but the operating-system lock does not. A new
owner can therefore acquire immediately and supersede stale data. Readers never
report stale data as current solely because the metadata file exists.

## Workbook Recovery Metadata

Recovery metadata represents an unresolved safety condition after xlflow can no
longer prove that Excel, VBA, COM automation, or bridge-side cleanup finished.
It is stored beside the lock and owner metadata as
`<LockID>.recovery.json`:

```json
{
  "schema_version": 1,
  "generation": "recovery-specific-random-token",
  "workbook": "C:\\projects\\sample\\sample.xlsm",
  "reason": "vba_may_still_be_running",
  "operation": "run",
  "xlflow_pid": 12345,
  "recorded_at": "2026-07-16T09:30:00Z",
  "session": {
    "active": true,
    "owner": "managed"
  },
  "excel_pid": 23456,
  "worker_pid": 34567
}
```

`generation` is created for each publication and is used for
compare-and-delete clearing. `workbook` must match the canonical identity
associated with the filename. `recorded_at` is UTC. `excel_pid`, `worker_pid`,
and session data are optional when the uncertain boundary could not observe
them. Suggested actions are derived at response time and are not persisted.

The file is diagnostic state and never an ownership lock. Its existence does
not make `busy` true, affect owner metadata, or prevent a process from acquiring
the normal operation byte range.

### Publication and Read Contract

Recovery publication occurs only for explicitly classified indeterminate
outcomes after Excel-side work may have begun. Examples include:

- `vba_may_still_be_running`;
- bridge or child-worker timeout during Excel/VBA execution or VBE compile;
- fatal COM/RPC failure after Excel activity began;
- failed or deliberately skipped cleanup that cannot prove Excel is idle;
- a poisoned session whose unsaved state requires discard; and
- outer provider timeout, abnormal termination, or invalid bridge response after
  the provider began workbook work.

Argument validation, source preflight, lock contention, ordinary cancellation
before Excel starts, and other failures known to occur before workbook work do
not publish recovery state.

The uncertain operation holds the normal workbook lease while atomically
replacing the recovery file. Publication happens before owner cleanup and lease
release. If publication fails, xlflow returns
`workbook_recovery_publication_failed`, retains the original uncertain error in
details, and gives emergency guidance to stop Excel manually.

Recovery reads and writes use the metadata publication guard. A command with
`recovery_behavior: block` acquires its normal lease, reads the marker, and
enters the command body only when no marker exists. A storage or read failure
returns `coordination_recovery_check_failed` and fails closed.

Malformed JSON, unsupported schema, filename/identity mismatch, or otherwise
untrustworthy metadata is represented as
`reason: "recovery_metadata_invalid"` and blocks unsafe operations. Invalid
metadata is not silently deleted.

### Recovery Diagnostic Contract

Blocked commands return operational exit code `3` with phase
`coordination.recovery`:

```json
{
  "status": "failed",
  "command": "push",
  "error": {
    "code": "workbook_recovery_required",
    "message": "The workbook is in an uncertain Excel state after a previous operation. Explicit recovery is required before this command can run; --wait will not resolve it.",
    "source": "xlflow",
    "phase": "coordination.recovery",
    "details": {
      "workbook": "C:\\projects\\sample\\sample.xlsm",
      "operation": "run",
      "reason": "vba_may_still_be_running",
      "recorded_at": "2026-07-16T09:30:00Z",
      "retryable": false,
      "wait_will_resolve": false,
      "recovery_actions": [
        "xlflow session stop --discard",
        "xlflow process cleanup 23456",
        "xlflow recovery clear",
        "xlflow recovery clear --force"
      ]
    }
  },
  "logs": []
}
```

The current command's operation is not substituted for the previous
`operation` that caused quarantine. Human output identifies the prior operation
when known, states that `--wait` does not resolve the condition, and recommends
only actions appropriate for the recorded session owner and process data.

An uncertain operation that successfully publishes its marker retains its
primary operation failure and adds top-level recovery data:

```json
{
  "recovery": {
    "required": true,
    "published": true,
    "reason": "vba_may_still_be_running",
    "operation": "run",
    "recorded_at": "2026-07-16T09:30:00Z",
    "excel_pid": 23456,
    "worker_pid": 34567,
    "session": {
      "active": true,
      "owner": "managed"
    }
  }
}
```

`excel_pid`, `worker_pid`, and `session` are omitted when unknown or inactive.

### Clearing Contract

Recovery state is cleared only while holding the normal workbook lease and only
when the generation still matches the marker that the recovery operation
examined.

- `recovery clear` is idempotent when no marker exists. With a marker, it clears
  only when `excel_pid` was recorded and the operating system confirms that PID
  no longer exists. A live, unknown, or unverifiable PID returns
  `workbook_recovery_verification_failed`.
- `recovery clear --force` removes the current generation without process
  verification and emits a mandatory safety warning. It does not stop Excel or
  VBA and does not repair or validate workbook contents.
- A managed `session stop --discard` clears the marker only after closing
  without saving and confirming the owned Excel process ended. Plain
  `session stop` and session save remain blocked during recovery.
- An external session stop detaches xlflow metadata but does not close the
  user's workbook or clear recovery state.
- `process cleanup <pid>` clears only matching markers for results with
  `terminated: true`.
- `process cleanup --auto` clears only markers for known PIDs that it actually
  terminated.
- `process cleanup --all` may also clear markers without an Excel PID only when
  cleanup completes and a new enumeration proves that no Excel processes remain.

Process cleanup enumerates candidate markers, acquires their workbook leases in
stable LockID order, performs cleanup, then generation-checks each clear. Partial
process failures leave the corresponding marker intact. Successful cleanup for
one workbook does not clear another marker merely because both were observed in
the same command.

## Recovery-Aware Command Start

A non-parallel-safe workbook operation follows this serialized sequence:

1. Resolve canonical workbook identities.
2. Acquire every required operation lease in stable LockID order.
3. Check recovery state for every target under the held leases.
4. If any target is blocked or unreadable, release all acquired leases in
   reverse order without entering the command body.
5. Run the command.
6. If the outcome is explicitly uncertain, publish recovery state before
   releasing the lease.
7. Clear matching owner metadata and release leases normally.

Different workbook identities remain independent. Direct Runner and CLI
invocations both follow this sequence. WSL delegates the full invocation and
uses the Windows-side state directory; it does not create a Linux marker.

## Busy Diagnostic Contract

Immediate contention returns `workbook_busy`. JSON output uses the standard CLI
error envelope and includes details sufficient for callers to decide whether to
retry:

```json
{
  "status": "failed",
  "command": "push",
  "error": {
    "code": "workbook_busy",
    "message": "Another xlflow operation is currently using this workbook.",
    "details": {
      "workbook": "C:\\projects\\sample\\sample.xlsm",
      "operation": "push",
      "resource_scope": "workbook",
      "retryable": true,
      "owner": null
    }
  }
}
```

When current metadata passes the guarded read, `owner` contains the metadata
fields suitable for diagnostics. Otherwise it is absent or null; missing owner
information never changes the `workbook_busy` result. Human output identifies
the workbook and operation as busy and indicates whether retrying is appropriate.
The CLI maps contention to environment/operational exit code `3`.

## Opt-in Waiting Contract

`--wait` opts a retryable, non-parallel-safe workbook command into lock waiting.
`--wait-timeout <duration>` overrides the finite 30-second default and is valid
only with `--wait`; zero and negative durations are invalid. Source-only,
parallel-safe, `excel_instance`, and non-retryable commands reject waiting from
their central policy before command execution.

The CLI uses one timeout for the complete sorted multi-workbook acquisition.
It releases partial acquisitions in reverse order on timeout, cancellation, or
error. The timeout ends when every required lease is acquired and never applies
to the command body. Acquisition first attempts the lock without waiting; only
actual contention enters polling. Human output prints one waiting line, while
JSON output remains a single envelope with no progress text.

Timeout returns `workbook_busy_timeout`; Ctrl+C or parent-context cancellation
returns `workbook_busy_cancelled`. Both use phase `coordination.acquire`, include
the attempted workbook, operation, resource scope, retryability, and configured
wait timeout, and map to operational exit code `3`. The underlying workbook
operation is never retried. Polling is cancellation-aware but does not guarantee
FIFO ordering.

Waiting ends when the operating-system lease is acquired. The command then
performs its recovery check once and returns `workbook_recovery_required`
immediately if quarantine is present. Recovery state is not contention, is not
retryable, and does not consume or extend the wait timeout.

## Status Observation

Top-level `status` and `session status` observe the configured workbook identity
through a combined coordination observation. They never infer ownership from
session metadata, owner metadata, or recovery metadata alone. The top-level
result always distinguishes active ownership from quarantine:

```json
{
  "coordination": {
    "busy": false,
    "recovery_required": true,
    "recovery": {
      "reason": "vba_may_still_be_running",
      "operation": "run",
      "recorded_at": "2026-07-16T09:30:00Z",
      "excel_pid": 23456
    }
  }
}
```

`busy` reflects only the OS operation lock. When current owner metadata is
available, the object also contains `resource_scope`, `operation_kind`,
`command`, `pid`, and RFC 3339 UTC `started_at`. `recovery_required` reflects
the marker independently; both fields can briefly be true while a failing
operation publishes quarantine or a recovery operation is running. Public
status omits marker generation, schema, internal filenames, worker PID, and
canonical-path storage details. It may include the recorded session owner when
the uncertain operation used a session.

The observation represents command-start state and may change before the
status response is returned. It is advisory and does not reserve the workbook;
later commands must still rely on normal CLI lock acquisition. If identity,
manager, or lock probing fails, session status preserves its bridge result,
omits `coordination`, and adds warning `coordination_status_unavailable`.
Recovery metadata failures are not reported as idle.

While recovery is required, status must not call unsafe workbook COM APIs. It
combines recovery metadata, project-local session metadata, and process liveness.
Unknown live fields remain unknown: session dirty state is null/unknown,
`source_of_truth` is `uncertain`, and `discard_required` is true. `process list`
similarly avoids workbook COM probes for recorded affected PIDs and returns
`has_workbook: null` plus `recovery_required: true`.

## Out of Scope for This Version

This contract does not add:

- FIFO queue behavior;
- an `excel_instance` lock primitive;
- a public policy/capabilities endpoint or serialized bridge schema; or
- a separate UserForm Designer lock outside workbook coordination.

Those features must consume this registry, identity, and lock contract rather
than creating parallel policy lists or workbook key algorithms.
