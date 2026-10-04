# ADR-0063: Canonical UserForm File Push and Picture Assets

## Status

Accepted

## Context

ADR-0060 established canonical FormSpec output for `pull --backend file`, while
ADR-0061 initially limited `push --backend file` to code updates on existing
UserForms. The pure-Go MS-OFORMS compiler can now generate forms and apply
supported Designer edits, so leaving file push dependent on compatibility
`.frm` / `.frx` artifacts would make pull output unusable as canonical source
and would keep asset changes outside source coordination and changed-only
decisions.

Pictures add a second authority boundary. Persisted MS-OFORMS data can retain
opaque picture bytes losslessly, but authored project assets need bounded
decoding, explicit replacement and removal, safe project-root resolution, and
deterministic file-pull output. Excel/VBIDE remains a separate authority and
must not be invoked as an implicit recovery path for an explicit file request.

## Decision

Supplement ADR-0060 and ADR-0061 with the following contracts:

- Canonical FormSpecs and their referenced picture assets are inputs to
  `push --backend file`. The command uses the existing pure-Go pack/compiler
  path and never launches or falls back to Excel, COM, VBIDE, or the bridge.
- File push treats the saved workbook as the authority for UserForm topology.
  Forms omitted from source remain in the workbook, and file push ignores
  `[pack].userform_topology`. `pack` keeps its existing independent topology
  setting and behavior.
- In sidecar code mode, a missing code sidecar preserves code in an existing
  form. An explicitly empty sidecar clears that code. A new form without a
  sidecar starts with empty code. `pack` keeps its existing code-source
  resolution contract.
- An omitted `picture` field preserves the current Image picture. A
  `picture: {path: ...}` value replaces it from a project-root-relative asset;
  `picture: {remove: true}` explicitly clears it. The field is limited to
  built-in Image controls and BMP/JPEG data. Null, empty, false removal,
  mixed path/removal, and unknown fields are rejected. Resolved asset bytes are
  runtime compiler input and are never serialized into YAML or JSON.
- Asset paths must resolve inside the project root after filesystem alias
  resolution, including symlinks and Windows junctions. Input assets are
  limited to 16 MiB and 16 million pixels. The enclosing Designer stream keeps
  its existing 64 MiB limit.
- File pull writes supported embedded pictures under the configured forms
  root's `assets/` directory using the full SHA-256 of the original image bytes
  and a format-derived `.bmp` or `.jpg` extension. Identical content may share
  one file. A conflicting existing content-addressed path fails before source
  publication. Picture assets are not automatically garbage-collected.
- Unsupported or unexportable existing picture resources remain in the saved
  workbook's Designer data. Pull does not invent asset paths; template-based
  operations preserve omitted resources, and an operation that would lose an
  unsupported resource fails before publication.
- The `compatibility_artifact_unsynchronized` marker describes stale or
  absent VBE export artifacts, not canonical FormSpec authority. When a
  canonical FormSpec and sidecar code are authoritative, stale `.frm` / `.frx`
  files are ignored as import inputs. `FRM201` continues to reject a marked
  `.frm` when `frm` code or compatibility import is authoritative.
- File-push changed-only fingerprints include canonical specs and referenced
  asset bytes. A state written by a backend that does not account for those
  inputs cannot prove a skip; cross-backend changed-only behavior is
  conservative and rebuilds when the effective input fingerprints differ.
- Custom ActiveX generation or compatibility import continues to require
  explicit Excel-authored `.frm` / `.frx` artifacts and an explicitly selected
  Excel path where supported. File operations fail with a capability error
  instead of falling back. The current real-Excel probe environment has no
  registered representative controls, so it does not establish custom-control
  import compatibility.
- Excel-backed `form build` and apply surfaces reject picture authoring until
  those surfaces implement the same validated picture contract.

## Consequences

- A file pull can produce canonical UserForm sources that a later file push
  can consume without requiring compatibility exports or Excel.
- Picture replacement and removal become reviewable source changes; asset-only
  edits invalidate a file-push skip.
- Existing template resources survive when the FormSpec omits `picture`, while
  an explicit removal remains distinguishable from omission.
- Content-addressed pull assets are stable and shareable, at the cost of
  retaining obsolete files until a user removes them deliberately.
- File push remains weaker than Excel for VBE compilation and arbitrary
  third-party ActiveX import. Lack of a registered custom-control fixture is
  an explicit unverified boundary.
- Backends can conservatively rebuild when their input sets differ, even when
  a prior push used the other backend.

## Alternatives Considered

1. **Treat every pulled `.frm` / `.frx` as authoritative.** Rejected because
   file pull does not regenerate those compatibility exports; importing a
   retained artifact could replace newer canonical Designer state.
2. **Delete forms omitted from file-push source.** Rejected because file push
   is an update of a saved workbook and must preserve existing topology. The
   separate `pack` topology setting remains the deliberate source-authority
   control for artifact generation.
3. **Interpret omission as picture removal.** Rejected because projected
   snapshots and partial edits routinely omit unsupported or unchanged state.
   Removal therefore requires an explicit value.
4. **Store image bytes inline in FormSpec.** Rejected because binary payloads
   would make source review, limits, path validation, and changed-only
   fingerprinting harder to reason about. Project-local assets keep the
   canonical spec small and readable.
5. **Fall back to Excel after a file-backend capability failure.** Rejected
   because it changes host requirements and mutation authority after the user
   explicitly selected file mode.

## Related

- ADR-0060
- ADR-0061
- Issue #912
- `docs/specs/file-push.md`
- `docs/specs/file-pull.md`
- `docs/specs/pack-command.md`
- `docs/specs/ms-oforms.md`
- `docs/specs/userform-picture-assets.md`
