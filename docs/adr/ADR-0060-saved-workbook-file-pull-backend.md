# ADR-0060: Saved-Workbook File Pull Backend

## Status

Accepted

## Context

`xlflow pull` historically exports VBA through Excel and VBIDE. That path is
required for complete UserForm Designer state, but it prevents agents and CI
from refreshing ordinary module source on hosts without Excel. The pure-Go
MS-OVBA reader used by `pack` can already parse saved `vbaProject.bin` state,
including its declared code page and module topology.

A second backend creates an authority choice. A saved workbook can differ from
an open, dirty Excel session, and a file reader cannot safely reproduce the
`.frm`/`.frx` export contract of VBIDE. Silently falling back between backends
would therefore hide both freshness and fidelity changes.

## Decision

Add `xlflow pull --backend file` while retaining `excel` as the default. The
backend is explicit and there is no automatic fallback in either direction.
The file backend supports saved `.xlsm` workbooks and reads only
`xl/vbaProject.bin`; it does not launch Excel, COM, the .NET bridge, or VBIDE.

The file backend exports standard, class, and document modules through the
same source-shape conventions consumed by `pack`. It decodes the project code
page, publishes UTF-8 source without BOM, applies configured folder annotations
and line-number removal, and transactionally reconciles only managed `.bas`
and `.cls` files. The configured forms tree is never reconciled by this
backend.

"Transactional" here means that the complete plan is validated before the
first mutation, each file replacement is atomic, and every affected managed
file is snapshotted so a returned publication failure attempts full rollback.
It does not mean crash-safe whole-tree atomicity. Abrupt process termination,
OS failure, or power loss can leave files from more than one generation. A
durable transaction journal was considered but deferred because restoring it
safely requires lifecycle, ownership, and damaged-journal policy beyond a
lightweight file-backend hardening change.

Before planning or publication, the file backend acquires source-tree leases
for the configured module, class, and document-module roots. The identities
are path-canonical and domain-separated from workbook identities. Each root is
exclusive while every canonical ancestor is held with a shared intent lease;
the complete set is acquired in stable LockID order on Windows and Unix. This
makes overlapping ancestor/descendant roots contend across projects without
serializing disjoint siblings. An explicit file pull does not acquire the
configured workbook lease because it never mutates workbook state.

If any UserForm is present, the complete pull fails before source mutation
with `pull_userform_unsupported`. Protected or malformed projects also fail
before mutation. A matching xlflow session does not change authority: the
saved file is read and `file_pull_live_session_ignored` warns that unsaved
live state was ignored. `--backend file` and `--session` are mutually
exclusive.

Capabilities schema version 2 publishes `default_backend` and per-backend
`requires_excel`. The command-level `requires_excel` remains `true` because it
describes the default command path. Consumers must use backend metadata when
planning an explicit non-default invocation.

## Consequences

- Ordinary VBA source can be refreshed deterministically without Excel on any
  host supported by the pure-Go reader.
- Callers can distinguish `saved_workbook` from live-session authority in JSON.
- Concurrent file pulls cannot reconcile the same managed source roots at the
  same time, while disjoint roots remain independent.
- A process crash releases source-tree ownership but does not roll back source
  files already published before the crash.
- Workbooks containing UserForms continue to require the Excel backend, even
  when only ordinary modules are of interest, so the source tree cannot become
  a partial mixed snapshot.
- `.xlam` and `.xlsb` remain available through the Excel backend but are not
  accepted by the initial file backend.

## Alternatives Considered

1. **Automatically fall back to Excel.** Rejected because backend selection
   changes host requirements and may change which workbook state is read.
2. **Export ordinary modules while skipping UserForms.** Rejected because a
   successful result would look complete while preserving stale form source.
3. **Block when a matching session record exists.** Rejected because explicit
   `--backend file` intentionally selects saved-file authority; a structured
   warning makes that boundary visible without making stale metadata fatal.
4. **Make the file backend the default.** Rejected because the existing Excel
   path supports more workbook formats and complete UserForm artifacts.
5. **Add a durable recovery journal now.** Deferred because a trustworthy
   automatic restore needs its own crash, corruption, and operator-recovery
   contract. The current guarantee remains explicit operation-level rollback
   for returned failures.

## Related

- Issue #871
- ADR-0012
- `docs/specs/file-pull.md`
- `docs/specs/workbook-coordination.md`
