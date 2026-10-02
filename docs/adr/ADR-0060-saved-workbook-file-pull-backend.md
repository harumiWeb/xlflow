# ADR-0060: Saved-Workbook File Pull Backend

## Status

Accepted

## Context

`xlflow pull` historically exports VBA through Excel and VBIDE. That path
prevents agents and CI from refreshing source on hosts without Excel. The
pure-Go project and MS-OFORMS readers can parse saved `vbaProject.bin` module
source and UserForm Designer state, including nested controls and the declared
project code page.

A second backend creates an authority choice. A saved workbook can differ from
an open, dirty Excel session, and a file reader cannot safely reproduce the
`.frm`/`.frx` export contract of VBIDE. Silently falling back between backends
would therefore hide both freshness and fidelity changes.

## Decision

Add `xlflow pull --backend file` and, after the file reader has passed its
production hardening gates, make `auto` the default. Explicit `file` and
`excel` requests never fall back. `auto` selects saved-file authority only
after proving that no matching live session or local Excel process owns the
workbook and that the file backend can produce a complete snapshot.
The file backend supports saved `.xlsm` workbooks and reads only
`xl/vbaProject.bin`; it does not launch Excel, COM, the .NET bridge, or VBIDE.

The file backend exports standard, class, and document modules through the
same source-shape conventions consumed by `pack`. In UserForm sidecar mode it
also projects Designer state to canonical `src/forms/specs/<Name>.yaml` and
publishes code-behind to `src/forms/code/<Name>.bas`. It decodes the project
code page, publishes UTF-8 source without BOM, applies configured folder
annotations and line-number removal, and transactionally reconciles all
managed source artifacts. Compatibility `.frm` / `.frx` artifacts remain
unmanaged because the pure-Go reader does not reproduce VBE exports.

"Transactional" here means that the complete plan is validated before the
first mutation, each file replacement is atomic, and every affected managed
file is snapshotted so a returned publication failure attempts full rollback.
It does not mean crash-safe whole-tree atomicity. Abrupt process termination,
OS failure, or power loss can leave files from more than one generation. A
durable transaction journal was considered but deferred because restoring it
safely requires lifecycle, ownership, and damaged-journal policy beyond a
lightweight file-backend hardening change.

Before planning or publication, the file backend acquires source-tree leases
for the configured module, class, and document-module roots, plus the form root
in sidecar mode. The identities are path-canonical and domain-separated from
workbook identities. Each root is exclusive while every canonical ancestor is
held with a shared intent lease; the complete set is acquired in stable LockID
order on Windows and Unix. This makes overlapping ancestor/descendant roots
contend across projects without serializing disjoint siblings. An explicit file
pull does not acquire the configured workbook lease because it never mutates
workbook state.

UserForms require `[userform].code_source = "sidecar"`. Compatibility `frm`
authority fails before source mutation with a dedicated capability error.
Form modules and Designer storages must match one-to-one; malformed,
unsupported, or inconsistent Designer state fails before publication. These
failures never fall back to Excel. A matching xlflow session does not change
authority: the saved file is read and `file_pull_live_session_ignored` warns
that unsaved live state was ignored. `--backend file` and `--session` are
mutually exclusive.

On Windows, automatic selection first honors `--session`, then validates a
matching session record against canonical workbook identity and the Excel PID
reported by Restart Manager. It registers only the configured workbook with
Restart Manager and identifies `EXCEL.EXE` without COM or external commands.
An open workbook, a valid live session, an indeterminate open-state probe,
compatibility `frm` UserForm authority, or an Excel-supported non-`.xlsm`
project selects the Excel backend. A closed sidecar UserForm project selects
the file backend only after its complete Designer snapshot can be parsed and
projected.
Automatic Excel selection retains workbook coordination, while automatic file
selection acquires only the source-tree leases used by the file backend. An
indeterminate open-state probe is attach-only: xlflow must find and attach to an
already-open matching workbook and must not open a second copy from saved state.
Malformed, protected, missing, unreadable, or unsafe saved projects retain
their deterministic validation failures rather than being hidden by an Excel
automation attempt. Non-Windows hosts use the file capability probe and fail
deterministically when it is unsupported; WSL remains local unless explicit or
validated live-session intent requires Windows delegation. When the WSL live
session probe itself fails, delegation preserves that failure as the same
attach-only selection and structured warning on Windows.

Successful output reports the actual backend, `auto` or `explicit` selection,
the stable selection reason, and source authority. Capabilities schema v3
publishes `default_backend=auto`; the auto backend keeps
`requires_excel=true` as a conservative advisory and adds
`selection=dynamic`.

Capabilities schema version 3 retains per-backend `requires_excel` and adds
dynamic-selection metadata. Consumers should use the selected backend reported
by command output rather than infer runtime authority from capability metadata.

## Consequences

- Ordinary closed-workbook VBA source is refreshed deterministically without
  paying Excel startup cost, while live or incomplete projects retain Excel
  fidelity.
- Callers can distinguish `saved_workbook` from live-session authority in JSON.
- Concurrent file pulls cannot reconcile the same managed source roots at the
  same time, while disjoint roots remain independent.
- An unrelated workbook lease does not block an automatic file pull, while an
  automatic Excel pull still participates in workbook coordination.
- A process crash releases source-tree ownership but does not roll back source
  files already published before the crash.
- Closed sidecar UserForm projects can be refreshed without Excel while still
  producing a complete canonical Designer/code snapshot. Compatibility
  `.frm` / `.frx` exports continue to require Excel.
- `.xlam` and `.xlsb` remain available through the Excel backend but are not
  accepted by the initial file backend.

## Alternatives Considered

1. **Silently fall back after an explicit backend request.** Rejected because
   it changes host requirements and source authority. Only `auto` owns the
   documented selection policy.
2. **Export ordinary modules while skipping UserForms.** Rejected because a
   successful result would look complete while preserving stale form source.
   The file backend instead validates and publishes the complete canonical
   sidecar snapshot in the same transaction.
3. **Block when a matching session record exists.** Rejected because explicit
   `--backend file` intentionally selects saved-file authority; a structured
   warning makes that boundary visible without making stale metadata fatal.
4. **Make `file` the fixed default.** Rejected because Excel remains necessary
   for open/live state, other workbook formats, and complete UserForm artifacts.
5. **Add a durable recovery journal now.** Deferred because a trustworthy
   automatic restore needs its own crash, corruption, and operator-recovery
   contract. The current guarantee remains explicit operation-level rollback
   for returned failures.

## Related

- Issue #880
- Issue #892
- Issue #871
- ADR-0012
- `docs/specs/file-pull.md`
- `docs/specs/workbook-coordination.md`
