# ADR-0061: Saved-Workbook File Push Backend

## Status

Accepted

## Context

`xlflow push` historically applies VBA source through Excel and VBIDE: the
bridge replaces non-document components, rewrites document-module code, and
saves the workbook. That path is required for new UserForm designer state and
for VBE compile validation, but it prevents agents and CI from applying source
changes on hosts without Excel — Linux runners, containers, WSL without a
Windows session, and remote development environments.

The pure-Go `pack` engine already reconstructs a complete `xl/vbaProject.bin`
from tracked source: it recompresses module streams, rewrites `PROJECT`/`dir`,
preserves opaque designer storage, and rejects protected, signed, or
structurally ambiguous projects. `coordination.PublishFile` already provides
validated atomic sibling replacement. A file push backend can therefore be a
composition of existing hardened pieces rather than a new mutation engine.

The authority boundary differs from file pull in one critical way: push writes
the saved workbook. A live Excel/VBE session may hold unsaved module state that
the source tree does not know about; overwriting the file underneath that
session would silently discard it. Where file pull can read a saved workbook
past a live session with a warning, file push must refuse.

## Decision

Add `xlflow push --backend file` as an explicit, opt-in backend. `--backend
excel` remains the default and the only backend that can create new UserForm
designer storage or claim VBE compile validation. Explicit backend requests
never fall back; `auto` is not offered for push.

The file backend composes existing pieces:

- Source enumeration and fingerprinting port `VbaSourceHelper.DiscoverSourceFiles`
  and `ComputeFingerprint` exactly, so `.xlflow/state/push.json` written by
  either backend is interchangeable and `--changed-only` interoperates across
  backends.
- Source transforms port `PrepareSourceForImport` semantics: document-module
  normalization, `'@Folder(...)'` annotation updates, Erl line-number
  instrumentation (with the same safety rejection of numeric labels and numeric
  `GoTo`/`GoSub`/`Resume` targets), and in-memory UserForm code sidecar merge.
- Workbook mutation is `pack.BuildWorkbook` applied to the saved file bytes;
  the staged artifact is structurally validated, then published by
  `coordination.PublishFile`'s atomic replace. The original workbook is never
  touched by preflight, backup, staging, or validation failure.
- `push.json` records `applied_to.saved_file` (path, last-write ticks, length)
  with the same schema the bridge writes, so a file push satisfies the bridge's
  `--changed-only` check and vice versa.

Safety gates are strict failures, not warnings: an Office `~$` lock file, a
matching `.xlflow/session.json` record (regardless of PID liveness, matching
the pack contract), a Restart Manager open-workbook report on Windows, an
indeterminate probe, or a matching WSL-side live session all abort before any
mutation with `push_workbook_open`/`push_active_session`. Source-tree leases
are acquired exclusively (shared on ancestors) exactly as in file pull. The
workbook lease is held only for the mutation window — template read through
atomic replace and `push.json` write — and under it the same recovery-marker
gate as other workbook mutators applies; a held lease fails fast with
`workbook_busy` because waiting there could deadlock against a command holding
the workbook lease while waiting on the source tree.

A successful file push means the artifact was reconstructed and structurally
validated — not that VBA compiled. Output reports `push.backend="file"`,
`push.target="saved_workbook"`, `vbe_validation="not_performed"`, and a
`vbe_validation_skipped` warning. Capabilities schema v3 keeps
`default_backend="excel"` and publishes `requires_excel=false` for the `file`
push backend. Backups record `backend="file"` in metadata.

## Consequences

- Agents and CI can apply source changes to a closed `.xlsm` without Excel,
  COM, VBIDE, or the bridge process.
- Push state is backend-agnostic: `--changed-only` skips correctly regardless
  of which backend performed the previous push.
- File push is strictly weaker than Excel push on UserForms: it can update
  code-behind on existing designer storage but cannot author new forms, so a
  source-only form fails with `push_userform_generation_unsupported`.
- A live session or open workbook is a hard block; users who intend
  saved-file authority must close Excel or stop the session first. This is the
  deliberate inverse of file pull's warn-and-continue.
- A crash between publish and `push.json` write leaves the workbook updated
  but the state stale; the next `--changed-only` conservatively re-pushes.
- `.xlsb`/`.xlam` and open-workbook pushes remain Excel-only.

## Alternatives Considered

1. **Extract a shared "push engine" consumed by both backends.** Rejected:
   the bridge's engine is fundamentally VBE-typed (VBComponents, CodeModule);
   `pack.BuildWorkbook` is already the pure-Go engine. Composition beats
   abstraction.
2. **Warn-and-continue past a live session like file pull.** Rejected: pull
   only reads; overwriting a file beneath an open dirty session destroys user
   work. Write authority must be stricter than read authority.
3. **Hold the workbook lease for the whole push.** Rejected: preflight,
   changed-only skips, and the rebuild must not contend on the workbook. The
   publish window still takes the lease non-blockingly — review of the
   merged PR showed skipping it entirely bypassed the recovery quarantine
   and let a concurrent rollback race the atomic replace.
4. **Offer `--backend auto` for push.** Deferred: auto-selection for a
   _writing_ operation needs a defensible policy for preferring saved-file
   over live Excel authority, which is a product decision rather than an
   implementation detail. Explicit-only keeps the contract honest.
5. **Reimplement fingerprint/state in a Go-only schema.** Rejected: sharing
   the bridge schema is what makes `--changed-only` and diagnostics portable
   across backends and across WSL/Windows delegation boundaries.

## Related

- Issue #895
- Issue #892, Issue #871
- ADR-0012, ADR-0016, ADR-0018, ADR-0020, ADR-0060
- `docs/specs/file-push.md`
- `docs/specs/file-pull.md`
- `docs/specs/workbook-coordination.md`
