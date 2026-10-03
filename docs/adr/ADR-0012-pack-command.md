# ADR-0012: Cross-Platform pack Command for Source-to-Artifact Workbook Generation

## Status

Accepted

## Context

xlflow treats VBA projects as source-controlled code. Most commands already operate on the source tree and are cross-platform. `push` is the one source-to-workbook command that still requires Windows and Excel: writing VBA back into a workbook means regenerating `xl/vbaProject.bin`, and today that happens through the Excel/VBIDE bridge as a live read-modify-write of the loaded workbook.

CI pipelines, containers, release packaging, and headless or agent-driven workflows need to produce an `.xlsm` artifact from source without a Windows-plus-Excel host. Everything else inside an `.xlsm` is OOXML XML that any platform can already produce. `xl/vbaProject.bin` — an OLE/CFB container holding MS-OVBA-compressed sources — is the single opaque piece that keeps that capability Windows-bound.

ADR-0008 already drew the relevant boundary. It placed source-tree management, lint/format/analysis, and release packaging in the Go core, and stated that source-only commands remain cross-platform while Windows/Excel automation stays in the bridge. Generating `xl/vbaProject.bin` from on-disk `.bas`/`.cls` sources is a source-to-artifact transform of the same kind: it works purely at the file level, never touches COM, VBIDE, or Win32, and never launches Excel.

A pure-Go, file-level writer cannot, however, provide what the live `push` path provides. It performs no VBE compile validation, no runtime validation, and gives no guarantee that every host-specific project state is interpreted exactly as Excel would. It is therefore not a drop-in replacement for `push`.

Feasibility is established. `ovba-writer` (https://github.com/kay-ws/ovba-writer), a standalone pure-Go reference implementation, reads and writes source-only `vbaProject.bin`, round-trips real Excel-compiled workbooks (cross-checked against `olevba`), and produces bins that real Excel recompiles and runs.

## Decision

Introduce a new, experimental, cross-platform command `xlflow pack` that builds an `.xlsm` artifact from the source tree, entirely in Go at the file level. Template mode regenerates `xl/vbaProject.bin` and replaces that one entry inside an existing workbook package. Explicit `--blank` mode authors both a minimal OOXML workbook and a fresh source-only VBA project without reading template bytes.

`pack` is a separate command, not a mode or backend of `push`:

- `push` is unchanged. It remains the Excel/VBIDE-backed live-session path and still requires Windows and Excel.
- The backend a command uses must not depend on project configuration. xlflow will not add an `xlflow.toml` switch that turns `push` into a file-level writer, because that would make a command's meaning depend on config and be confusing in CI logs, in support and debugging, and for agents.

The initial MVP boundary is deliberately narrow and fail-loud:

- the command is gated behind `--experimental`;
- `--out` is required; `pack` never overwrites the template or configured source workbook in place;
- `pack` operates only on closed workbook files, never on active sessions or live workbooks;
- `--template` is optional and falls back to the source workbook configured in `xlflow.toml`;
- `--blank` is mutually exclusive with an explicit `--template`; when neither is supplied, the configured source workbook remains the legacy template;
- blank mode deliberately fixes host topology to `ThisWorkbook` plus one worksheet/code name `Sheet1`; both document sources must exist, additional document modules and all UserForms fail loudly;
- blank mode authors deterministic VBA/Office reference records, takes its project code page from `[pack.blank]` (default `1252`), and uses the canonical MS-OVBA LCID `0x00000409` rather than inferring locale from a host or template;
- standard and class modules are supported first; document modules only where they map safely against the template;
- JSON output identifies the backend as the experimental pure-Go packer and reports that VBE compile validation was not performed;
- the pure-Go packaging path has Linux tests;
- unsupported cases fail with specific errors rather than best-effort behavior.

Unsupported in the MVP, each a loud and specific error: active sessions or live workbooks, in-place overwrite of the template/source workbook, protected VBA projects, signed VBA projects, full UserForm/`.frx` generation, and unknown or ambiguous VBA project layouts.

The contract for command shape, the JSON envelope, and exit codes lives in `docs/specs/pack-command.md`.

### Ownership and attribution

`pack`'s implementation, test harness, and compatibility behavior live in an xlflow-internal package, `internal/pack`. xlflow owns the public behavior and the maintenance responsibility, because this is a core artifact-generation path.

`ovba-writer` is a reference implementation and feasibility proof, not a long-term dependency. xlflow does not take a direct module dependency on it; the logic is reimplemented in-repo and the reference is credited by link. Both `ovba-writer` and xlflow are MIT-licensed, so reuse is license-compatible. Because `pack` is not a `go.mod` dependency, no entry in `THIRD_PARTY_LICENCES.md` is required.

### Staged scope

`pack` matures before it loses the experimental gate:

- **Initial MVP** — as described above.
- **Before non-experimental status** — broader document-module fixtures; hardened signed/protected project detection beyond the MVP's baseline reject-on-detect; non-ASCII/Japanese source fixtures; Windows/Excel smoke tests that open the generated workbook and compile/run a minimal macro; a documented UserForm preservation/update strategy; and docs stating that `pack` does not compile or run VBA.
- **UserForms, staged in three steps** — (1) preserve the template's existing designer streams unchanged, generating no forms; (2) update form code-behind while keeping the template's designer state; (3) full reconstruction from exported `.frm`/`.frx` as a separate, higher-risk phase. Only step (1)'s "carry existing designer streams through untouched" is compatible with the MVP, and only when the forms already exist in the template.

### Amendment: canonical UserForm generation foundation (Issue #883)

The Stage-3 foundation generates new Designer state from the canonical
`xlflow.userform` FormSpec, independently of exported `.frm` / `.frx`
compatibility artifacts. The lossless MS-OFORMS model remains separate from
the authoring schema. Generation returns a reparsed, signed model; it does
not weaken the no-op serializer's protection against arbitrary model edits.

Root client dimensions are explicit `form.build.clientWidth` / `clientHeight`
in points. Existing width/height fields retain their outer-dimension contract.
Using an assumed window-chrome offset would make generation depend on Excel
environment and cause repeated round trips to grow forms, so this stage
does not convert outer dimensions. Excel-backed Designer build rejects the
new client input before mutation until it has a supported application path.

Issue #883 owns flat common-control generation and code-behind/project
assembly with an already-present Microsoft Forms reference. Reference
mutation remains Issue #886; CLI source-authority planning and blank/template
pack integration remain Issue #887. The current CLI restrictions therefore
remain in force. Deterministic Designer bytes use only fixed class IDs;
cryptographically random GUIDs are limited to new component identity in
code-behind, with injectable generation for reproducible tests.

### Amendment: first-Forms reference and blank UserForms (Issue #886)

Blank pack now creates the supported flat UserForm subset from canonical
FormSpec and configured code authority. Compatibility `.frm`/`.frx` Designer
bytes are not generation inputs, because making them authoritative would
conflict with file-pull's canonical specs. Template Designer integration
remains Issue #887.

Reference mutation uses complete MS-OVBA reference groups, retaining existing
records verbatim. Missing Forms references are appended as a canonical
REGISTERED Forms 2.0 reference using the existing fixed Windows path policy
for blank references. This avoids persisting a developer's user-specific
extended-type-library cache path. Existing CONTROL references remain intact
and are identified by OriginalTypeLib. Excel release-gate evidence verifies
actual resolution, compilation and save/reopen rather than assuming REGISTERED
and CONTROL forms are operationally interchangeable.

### Amendment: template UserForm source planning (Issue #887)

Template pack now generates supported forms from canonical specs and applies
supported property edits to existing lossless Designers. Compatibility
`.frm`/`.frx` Designer bytes remain generated artifacts; `.frx` is not needed
to author a supported form. Legacy `.frm`-only existing forms retain code-only
updates, preserving compatibility with earlier pack workflows.

`[pack].userform_topology` defaults to `template`, retaining omitted forms.
Explicit `source` authority removes omitted forms, including the final form,
and requires canonical specs for all supplied forms. Deletion is opt-in because
partial source trees have historically preserved template forms; inferring
authority from the presence of a spec would silently delete unrelated forms
and could not represent an intentionally empty form set.

An existing spec lists the complete control topology, while omitted properties
retain binary values. Supported edits preserve opaque records and existing
component identities. Topology changes outside the supported Frame/common-
control boundary and unsupported property changes reject the whole plan; pack
never rebuilds an existing Designer to bypass those restrictions.
Removal reconciles module, Designer, PROJECT declarations, Workspace entries,
and PROJECTwm while retaining all references that other code may still use.

### Amendment: Frame topology compilation (Issue #884)

Frame and common-control topology now follows a complete canonical control
list: additions, deletions, replacements, reparenting and sibling ordering.
Existing controls are matched by exact, case-sensitive name and
case-insensitive type; source IDs express hierarchy rather than binary
persistence identity. Retained binary records,
resources and storage metadata are reused instead of regenerating the entire
Designer, since the authoring schema cannot represent every persisted field.
Unknown container bookkeeping rejects structural changes rather than risking
its loss. MultiPage/Page/TabStrip topology requires its separate implementation.

Coordinates always remain relative to the owning parent. An omitted position
keeps its numeric value after reparenting; there is no implicit absolute-position
conversion. This keeps source intent deterministic without depending on host
window chrome, nested client offsets or display DPI. Structural changes follow
the same clone, encode, reparse and read-back validation boundary as property
compilation, while the ordinary serializer continues to reject arbitrary model
mutation.
Writer output is reparsed and compared with the planned project before atomic
publication. This adds structural validation without changing the permanent
no-VBE-validation contract. Issue #884 supersedes the earlier Frame-generation
follow-up for Frame/common-control hierarchies; MultiPage/Page/TabStrip editing
remains a separate follow-up.

The Excel gate also established that inspecting root Caption through Designer
does not establish runtime behavior: instantiated forms use the VBFrame text
value. Generation and explicit root-caption edits therefore keep FormControl
and VBFrame values consistent. Values unrepresentable in the project code page
are rejected rather than silently changing the runtime caption. This restriction
does not apply to control captions or values, which retain Unicode persistence.

### Amendment: stable graduation (Issue #858)

The pre-stable stages are complete. `pack` is now a stable command and no
longer requires `--experimental`. The first stable release accepts that flag as
a hidden deprecated no-op so existing automation can migrate without changing
artifact behavior; it is not part of the stable help, JSON, or error contract.

At stable graduation, the authority boundary was as follows. The UserForm
restrictions in this historical boundary are superseded by the canonical
blank-generation and template-source-planning amendments above (Issues #883
and #887).

- standard/class component topology and code are source-authoritative;
- document topology is template-authoritative and matched document code is
  source-authoritative;
- existing UserForm topology and designer state are template-authoritative,
  while code-behind follows `[userform].code_source`;
- creating a new UserForm remains unsupported and fail-loud.

Graduation is backed by the manual Windows/Excel release gate. It exercises
standard/class add, update, remove, and rename; multiple document modules and
CodeNames; non-ASCII source and component names; existing nested UserForm and
`.frx` state; references and project metadata; and compile/run sentinel checks.
This evidence does not change the runtime contract: ordinary `pack` executions
remain Excel-independent and report `vbe_validation = "not_performed"`.

### No VBE validation

"No VBE validation" is a permanent semantic boundary, not a temporary limitation. `pack` never compiles or executes VBA. Its output is a file artifact whose correctness against the VBE has not been verified; Excel compiles it from source on first open. Every `pack` run reports this in its output. Consumers that need compile or runtime validation must use the Excel/VBIDE-backed `push` path on Windows.

### Amendment: canonical blank-project identity metadata (Issue #870)

Blank generation exposes only `code_page` as locale-sensitive configuration.
MS-OVBA requires both `PROJECTLCID` and `PROJECTLCIDINVOKE` to be
`0x00000409`, independently of `PROJECTCODEPAGE`, so `pack.blank.lcid` is
removed and rejected as an unknown pack key. CP932 remains the supported way
to encode Japanese project text and component names.

The deterministic Project CLSID and its matching CMG/DPB/GC tuple remain
fixed. The Windows/Excel release gate opens two independently generated blank
workbooks concurrently and runs both macros. It records unbroken VBA, Excel,
stdole, and Office references; a project-to-project reference succeeds after
one default `VBAProject` name is changed. The observed ambiguity between two
unchanged projects is therefore the normal duplicate project-name boundary,
not a CLSID collision, and does not justify sacrificing byte-for-byte project
metadata determinism.

### CFB and OVBA compatibility boundary

The internal artifact path supports both CFB v3 and v4 and preserves the
template's major version. It implements DIFAT output instead of rewriting large
or v4 templates into a smaller v3-only profile. This keeps the template as the
authority for its storage geometry and avoids an implicit format migration.
The repository's current Excel-derived fixture survey found v3 containers, but
that evidence is not treated as a stable guarantee that user templates cannot
use v4; both geometries are therefore part of the supported boundary.

Because CFB and MS-OVBA inputs can contain attacker-controlled lengths and
links, parsing is fail-loud and bounded. Sector and directory graphs reject
cycles, invalid references, and shared ownership; MS-OVBA decompression has a
64 MiB per-stream output limit. Native fuzz targets are maintained alongside
focused regression fixtures. A malformed or unsupported container remains a
content validation failure and never produces a best-effort artifact.

## Consequences

- Positive: `push` semantics are unchanged; the live-session workflow does not regress.
- Positive: a cross-platform path exists for CI, containers, release packaging, and headless or agent artifact generation, with no Windows-plus-Excel host.
- Positive: keeping the backend independent of configuration means a command's meaning stays stable and legible in CI logs and to agents.
- Positive: ownership in `internal/pack` keeps the artifact-generation path's maintenance and compatibility behavior inside xlflow.
- Positive: explicit blank mode removes the last workbook-template dependency for the narrow one-sheet release profile without changing the legacy template fallback.
- Positive: preserving CFB v3/v4 geometry and supporting DIFAT avoids silent
  template format conversion and removes the header-only FAT size ceiling.
- Negative: `pack` output lacks VBE compile and runtime validation; consumers must treat it as unvalidated, and the contract must keep saying so.
- Negative: two commands (`pack`, `push`) with overlapping inputs but different guarantees increase the surface to document and explain.
- Negative: xlflow takes on maintenance of MS-OVBA and OLE/CFB compatibility behavior in-repo rather than delegating it to an external module.
- Negative: supporting both CFB geometries and bounded graph parsing adds more
  compatibility and adversarial-input cases to the internal implementation.
- Negative: stable compatibility still depends on maintaining pure-Go fixtures
  plus a manual Windows/Excel release gate; PR CI alone cannot prove VBE
  compile/runtime compatibility.
- Negative: blank mode is intentionally not a general workbook-layout generator;
  projects needing sheets beyond `Sheet1` still require a template. UserForm
  generation is limited to the supported canonical-spec subset.

## Alternatives Considered

1. **Make `push` itself cross-platform / Excel-free** — Rejected because it would change `push`'s validation guarantees and conflate the live-session, VBE-validated semantics with an unvalidated file-level transform under one command name.
2. **Add an `xlflow.toml` switch that changes `push`'s backend** — Rejected because a command's behavior would then depend on project configuration, which is confusing in CI logs, in support and debugging, and for agents.
3. **Implement Excel automation in Go** (ADR-0008 Alternative #2) — Not applicable and rejected. `pack` never touches COM, VBIDE, or Win32 and never launches Excel. It is a file-level transform, not automation, so it does not recreate interop complexity in the CLI process.
4. **Take a long-term direct dependency on `ovba-writer`** — Rejected because a core artifact-generation path should not depend on an external module whose maintenance xlflow cannot control. The implementation is internalized under `internal/pack`, with the reference credited.
5. **Include full UserForm/`.frx` generation in the MVP** — Rejected as higher-risk. UserForm support is staged, starting with carrying existing template designer streams through untouched.

## Related

- `docs/adr/ADR-0008-dotnet-excel-bridge.md`
- `docs/adr/ADR-0001-agent-ready-vba-cli-architecture.md`
- `docs/specs/pack-command.md`
- `docs/specs/cli-contract.md`
- https://github.com/kay-ws/ovba-writer
- xlflow issue #143
