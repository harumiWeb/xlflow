# UserForm Picture Assets

This specification defines the external image input and extraction contract
for the built-in MSForms `Image` control. The pure-Go compiler persists
supported image data into the UserForm Designer; the workbook does not need
the source asset at runtime. The compiler and file-pull paths do not launch or
fall back to Excel.

## FormSpec field

`picture` is optional and valid only on an `Image` control:

```yaml
- id: brand-mark
  name: BrandMark
  type: Image
  picture:
    path: src/forms/assets/brand.jpg
```

The path is relative to the xlflow project root, independent of the caller's
current directory. It must resolve to a file inside that root after resolving
filesystem aliases, including symlinks and Windows junctions. A path that
escapes the project root is rejected before workbook mutation or source
publication.

Only BMP and JPEG assets are accepted. The file content must be a valid image,
and the image is limited to 16 MiB and 16 million pixels. The full
MS-OFORMS Designer stream remains subject to its 64 MiB per-stream limit.
Accepted BMP input uses a 40-byte `BITMAPINFOHEADER`, uncompressed `BI_RGB`,
and 24-bit or 32-bit pixels. JPEG content is fully decoded during validation.

The `picture` value has exactly one of these forms:

```yaml
picture:
  path: src/forms/assets/brand.bmp
```

```yaml
picture:
  remove: true
```

Omitting `picture` preserves an existing picture. `path` replaces the picture;
`remove: true` clears it. For a newly generated Image control, omission means
there is no picture. Null, empty values, `remove: false`, a value containing
both `path` and `remove`, unknown keys, or `picture` on another control type
are invalid. Rejected input produces no partial Designer or published
workbook.

Decoded bytes are held in the compiler's in-memory `FormSpecPicture.Data`
field. That field is excluded from YAML and JSON serialization; project source
contains only the relative asset path or explicit removal intent.

## File push and pack

`push --backend file` passes canonical FormSpecs and resolved assets through
the shared pure-Go UserForm compiler. Both FormSpec files and referenced asset
bytes participate in file-push coordination and changed-only fingerprints.
An asset-only edit therefore rebuilds the workbook. A state fingerprint from a
backend that did not include these inputs cannot justify a skip.

File push keeps saved-workbook UserForm topology: omitted forms remain, and
`[pack].userform_topology` does not affect file push. In sidecar mode, a missing
code sidecar preserves existing code in an existing form; an empty sidecar
clears it. A new form with no code sidecar has empty code. `pack` retains its
own documented topology and code-source rules.

Only image assets referenced by canonical FormSpecs are loaded, validated, and
included in file-push fingerprints. Unreferenced files under the reserved
forms `assets/` directory are retained and ignored. Pull never garbage-collects
assets; deleting obsolete files is optional.

Picture authority follows the FormSpec reference rather than a filename suffix.
A referenced `.bas` image, including a project-contained relative symlink,
must not become VBA module source. Fingerprints retain the logical reference
path and captured image bytes; file reads and target-directory coordination
use the resolved physical path. Regression coverage verifies that a symlink
target edit invalidates changed-only state.

Template-based pack and file-push operations preserve a picture when the
FormSpec omits `picture`. Explicit path or removal intent is applied to the
Image control. Blank pack can generate a supported picture from its asset.
Excel-backed `form build` and apply APIs reject picture authoring until they
implement this contract. An explicit file backend reports unsupported cases
without starting Excel.

Custom ActiveX controls and arbitrary non-Image picture properties are outside
this contract. File operations fail before publication when completing the
requested operation would require unsupported custom-control authoring or
lossy resource regeneration.

## File pull

File pull extracts a picture only when it can validate and decode the stored
BMP/JPEG payload within the same limits. The emitted path is content-addressed
under the configured forms root:

```text
<forms-root>/assets/<full-lowercase-SHA-256>.bmp
<forms-root>/assets/<full-lowercase-SHA-256>.jpg
```

The hash is computed over the decoded image bytes written to the asset file,
not over the containing StdPicture wrapper. Excel can normalize an authored
JPEG to BMP before pull, so the resulting BMP path is hashed from those BMP
bytes. Identical content may be referenced by multiple controls and shares one
file. If a planned path already exists with different bytes, pull fails before
publishing any source changes. Pulled picture assets are retained; pull never
garbage-collects old assets.

Unsupported or unexportable embedded picture data remains in the saved
workbook's Designer storage. Pull does not create a guessed or fabricated
asset path. A later template-based operation preserves the omitted picture;
any operation that would discard or regenerate unsupported data fails before
publication.

The `compatibility_artifact_unsynchronized` warning describes VBE `.frm` /
`.frx` exports and does not make canonical FormSpec sidecar input unsafe. Stale
compatibility files are not import authority when a canonical spec and
sidecar code are selected. The marker still blocks operations that select
`.frm` code or compatibility artifacts as import authority.

## Excel evidence

The developer-only `scripts/test-userform-pictures-e2e.ps1` Excel-authored
fixture loads 24 x 16 pixel BMP and JPEG files into Image controls, saves and
reopens the workbook, and checks picture type, dimensions, and a VBA sentinel.
The committed capture records Excel 16.0 build 17932 on Windows 10.0.22631,
with confirmed Excel-process cleanup. Both input formats were persisted by
Excel as a StdPicture resource with GUID
`0452e30b918fce119de300aa004bb851`, preamble `6c740000`, payload length
`0x4b6`, and a payload beginning with `BM`. Excel therefore normalized the
authored JPEG input to BMP in this fixture. Each image was observed as 508 x
339 HIMETRIC units before save and after reopening.

This fixture confirms Excel's LoadPicture and save/reopen path for BMP and
JPEG inputs, plus persisted BMP resource handling. It is separate from the
pure-Go direct-JPEG check: the authored JPEG is normalized to BMP before this
fixture records the saved resource.

The full canonical Issue #912 picture gate was then verified in real Excel
with source asset paths unavailable during workbook verification. The
file-push, blank-pack, and template-pack outputs all displayed both Image
controls successfully. The native JPEG StdPicture representation passed in
the file-push output and in both blank and template pack. On Excel 16.0 build
17932 / Windows 10.0.22631, both controls reported StdPicture type 1, positive
dimensions of 508 x 339 HIMETRIC, picture size mode 3, and alignment 3 before
save and after SaveAs/reopen. The VBA sentinel was `issue-912-ok` and owned
Excel-process cleanup was confirmed for all three verification runs.

Evidence is retained in these relative workspaces:

| Workflow            | Gate workspace                                              | Excel verification workspace                               |
| ------------------- | ----------------------------------------------------------- | ---------------------------------------------------------- |
| Full pull/file-push | `tmp_workspaces/issue-912-file-push-20261004-040504-de315a` | `tmp_workspaces/issue-912-pictures-20261004-040532-dc56f3` |
| Blank pack          | same gate workspace                                         | `tmp_workspaces/issue-912-pictures-20261004-040537-73f28c` |
| Template pack       | same gate workspace                                         | `tmp_workspaces/issue-912-pictures-20261004-040541-6d63b5` |

This establishes Excel display and save/reopen compatibility for native JPEG
picture authoring through file push, blank pack, and template pack in that
environment. It does not establish arbitrary ActiveX compatibility or
compatibility with every Excel version. When Excel normalizes an authored
JPEG to BMP on pull, its extracted content-addressed asset is BMP, so byte
equality with the original JPEG is not expected. A no-op/template carry path
must retain the original persisted picture resource identity. Re-encoding is
assessed by equivalent pixels and Designer behavior rather than equality with
the originally authored JPEG bytes.

## Related

- ADR-0063
- `docs/specs/file-push.md`
- `docs/specs/file-pull.md`
- `docs/specs/pack-command.md`
- `docs/specs/ms-oforms.md`
