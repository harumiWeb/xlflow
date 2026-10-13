# ADR-0072: Canonical UserForm Compilation and Persisted Root Dimensions

## Status

Accepted

## Supersession

This decision supersedes only the UserForm root-dimension and backend-split
statements in the Issue #883 amendment to ADR-0012, and the final
Excel-backed picture-authoring restriction in ADR-0063. The remaining pack,
file-push, and picture-asset decisions in those records continue to apply.

## Context

The unreleased FormSpec contract exposed the persisted root `DisplayedSize` as
`form.build.clientWidth` / `clientHeight`, while older `form.width` / `height`
fields described a different, outer Designer size. Excel-backed build and apply
then rejected the persisted-size fields, and template pack and file push
rejected root-size edits. Those separate meanings made a single FormSpec behave
differently across commands and encouraged callers to estimate a fixed Excel
window-chrome offset.

Picture assets already use the pure-Go MS-OFORMS compiler for pack and file
push, while Excel-backed FormSpec build and apply have a separate bridge path.
That split also makes picture support and state-preservation rules vary by
command. Canonical FormSpecs should have one validation and compilation
contract; Excel automation should transfer the compiler-produced form artifact
into the target workbook while preserving the target workbook as the baseline.

Canonical FormSpecs are intended to author Designer state without requiring
prebuilt `.frm` / `.frx` exports. Those files remain useful compatibility and
pull/build artifacts. Code-behind authority remains governed by
`[userform].code_source` and the existing sidecar rules.

## Decision

- `form.width` and `form.height` are the canonical root dimensions. They are
  points describing the persisted MS-OFORMS `DisplayedSize`. Snapshot and pull
  project that value into the inspection `client_width` / `client_height`
  fields; those names do not identify VBFrame `ClientWidth` / `ClientHeight`
  as the source. VBFrame client dimensions are compiler-maintained companion
  values and never seed root authoring intent. Excel import/export may
  normalize `DisplayedSize`, so snapshots report the actual persisted value
  and exact round-trip equality with authored points is not guaranteed. This
  Excel behavior is separate from the pure-Go compiler's nearest-integer
  HIMETRIC conversion and rounding contract. A newly generated form defaults
  omitted axes to 240 by 180 points; when editing an existing template form,
  omitted axes preserve the baseline dimensions. Captured persisted DisplayedSize
  seeds the direct root fields when a snapshot is projected, while observed-only
  fields are never used as a fallback. Root `form.build` aliases, including
  `build.caption`, `build.width`, `build.height`, `build.clientWidth`, and
  `build.clientHeight`, are invalid fields. Snapshot-only `form.observed` data
  never supplies authoring values.
- One pure-Go UserForm compiler validates and compiles FormSpecs for
  `form build`, `form apply`, Excel-backed `push`, file `push`, and `pack`.
  Picture asset resolution, validation, and compilation use that same compiler
  contract. For Excel-backed `form build` and `form apply`, the bridge first
  saves a private copy of the live target workbook with `SaveCopyAs`. Go then
  runs `pack.BuildWorkbook` against that baseline and the target FormSpec. The
  bridge opens the compiled temporary workbook in a private Excel instance,
  exports the compiled form to private `.frm` / companion `.frx` artifacts,
  and imports that component into the target workbook through
  `VBComponents.Import`. If replacement fails, the bridge removes a partial
  import and restores the prior form from its exported backup. There is no
  imperative Designer-property or geometry-write fallback: root dimensions and
  picture data arrive in the compiler-produced artifact.
- Compilation and application preserve existing code-behind and unsupported
  or opaque Designer state unless the requested operation explicitly changes
  supported state. If an operation cannot preserve data it does not understand,
  it fails before publishing or mutating the workbook. FormSpec Designer
  authoring does not require prebuilt `.frm` / `.frx` inputs; successful
  workbook-backed operations may continue to export those compatibility
  artifacts.
- `pack --check` runs the same source preflight, source plan, FormSpec and
  picture compilation, VBA-project generation, and structural readback as
  ordinary `pack`. It skips publication and writes no artifact.
- `new` and `init` seed `[pack.blank].code_page` from the actual configured
  workbook's VBA project `PROJECTCODEPAGE`. Managed VBA and FormSpec sources
  remain UTF-8 without BOM and are not transcoded by scaffolding or packing.
  A valid UTF-8 source character that cannot be represented in the target
  project code page produces `pack_source_encoding_failed` with path, line,
  Unicode-character column, rune, rune code, and code page. Multiple pack
  validation failures aggregate under `pack_validation_failed`, with each
  issue retaining its specific compiler code. Invalid UTF-8 remains the
  separate `source_encoding_invalid` preflight diagnostic with byte-position
  details; both are reported before compilation.
- `init` defaults imported projects to `[userform].code_source = "sidecar"`.
  The explicit `--userform-code-source frm` option remains available for
  artifact imports that require embedded `.frm` code as the authority,
  including custom-control workflows.
- Designer snapshots read `DisplayedSize` from the UserForm in a private
  `SaveCopyAs` workbook copy and expose the value as `client_width` /
  `client_height`. They do not derive these values from VBFrame client
  properties, outer Designer size, or runtime `InsideWidth` / `InsideHeight`.
  Excel may normalize `DisplayedSize` during import/export; a snapshot records
  the persisted result rather than promising exact equality with the original
  authored value. The pure-Go HIMETRIC conversion contract is independent of
  this Excel normalization.

## Consequences

- FormSpec root geometry has one meaning across authoring, preview, file
  generation, and Excel-backed application. This removes the old outer/client
  dimension conflict and prevents drift from guessed offsets.
- The compiler becomes the authority for supported Designer and picture
  semantics. Excel remains required to mutate a live workbook and to establish
  VBE/runtime behavior; a successful compile or `pack --check` does not prove
  those runtime properties.
- Existing source that uses `form.build.*` must be migrated to the direct
  `form` fields before this unreleased contract ships. Strict validation
  intentionally rejects those aliases instead of maintaining two schemas.
- Removing the `.frm` prerequisite simplifies canonical source, while
  `[userform].code_source = "frm"` may still select a `.frm` as the code source.
  Omitted code and unsupported Designer state retain their established
  preservation semantics.
- `pack --check` can catch source and artifact-structure failures without
  creating an output workbook, but it does not add VBE compile or runtime
  evidence.
- Scaffolding records the workbook's actual project code page without changing
  UTF-8 sources. Projects with an unsupported or unreadable VBA project fail
  scaffolding rather than receiving a guessed code page.

## Alternatives Considered

1. **Keep outer and client dimension families.** Rejected because callers
   would still need backend-specific field selection, and Excel's outer size
   depends on host presentation details.
2. **Convert between outer and client size with a fixed offset.** Rejected
   because the offset is not stable across Excel versions and layouts; it also
   changes persisted size across repeated round trips.
3. **Keep separate Excel and pure-Go FormSpec compilers.** Rejected because
   validation, pictures, unsupported-state handling, and root geometry could
   diverge again.
4. **Require `.frm` / `.frx` as the generation input.** Rejected because
   exported artifacts are not the canonical Designer source and are
   unnecessary when a FormSpec and its referenced assets fully describe the
   requested supported state.
5. **Make `pack --check` a partial source-only check.** Rejected because it
   would not establish that the normal compiler and VBA-project writer can
   produce a structurally valid artifact from the same inputs.

## Related

- [ADR-0012](ADR-0012-pack-command.md)
- [ADR-0063](ADR-0063-canonical-userform-file-push-and-picture-assets.md)
- [`ms-oforms.md`](../specs/ms-oforms.md)
- [`pack-command.md`](../specs/pack-command.md)
- [`file-push.md`](../specs/file-push.md)
- [`userform-picture-assets.md`](../specs/userform-picture-assets.md)
- [`cli-contract.md`](../specs/cli-contract.md)
