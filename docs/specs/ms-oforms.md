# MS-OFORMS persistence and projection contract

This document defines xlflow's internal, pure-Go persistence contract for
Microsoft Forms designer state stored in `vbaProject.bin`. The architectural rationale
for file-level VBA project handling remains in ADR-0012. The user-facing
UserForm schema remains the `xlflow.userform` `FormSpec`; the model described
here is a separate binary persistence model.

## Scope and entry point

`internal/vba/userforms/oforms.ReadForm` reads one top-level UserForm storage
from an already-open `internal/pack/cfb.Container`. Callers supply the storage
path and the VBA project's `PROJECTCODEPAGE`. The reader does not use Excel,
COM, VBIDE, or Windows APIs and does not silently fall back to an Excel-backed
path.

The reader owns the following structures:

- the top-level and nested `f` and `o` streams;
- `FormControl`, `FormSiteData`, `SiteDepthsAndTypes`, and
  `OleSiteConcreteControl` records;
- property masks, aligned DataBlock and ExtraDataBlock fields, StreamData,
  `TextProps`, geometry, names, IDs, and object-stream extents;
- `\x03VBFrame`, `\x01CompObj`, optional `x` streams, CFB storage metadata,
  and nested container storages;
- raw strings, padding, pictures, arrays, record tails, and streams required by
  the lossless serializer and later controlled mutation stages.

`DiscoverForms` reports root-level non-`VBA` storages containing an `f` stream.
Module classification remains the responsibility of the VBA project layer.
`SerializeForm` re-emits one model accepted by `ReadForm` as a validated set of
full-path streams and storage metadata. It performs no Excel, COM, or VBIDE
fallback.

## Lossless binary model

Decoded values never replace their persistence bytes. Strings retain their
compression flag and original encoded bytes. Records retain masks, alignment
padding, opaque pictures and arrays, raw tails, and the complete source record.
Each container level retains its original streams and storage metadata.
Known Designer stream identities follow MS-CFB case-insensitive name matching;
the model also retains and replays the directory entry's original spelling.
Case variants such as `F`, `X`, or `\x01COMPOBJ` therefore remain byte-for-byte
present instead of being normalized or dropped.

The first known property tables cover the built-in classes represented by the
existing FormSpec contract: Label, TextBox, ComboBox, ListBox, CommandButton,
CheckBox, OptionButton, and Frame. The shared reader also recognizes Image,
MorphData, SpinButton, TabStrip, ToggleButton, ScrollBar, Form/Page, and
MultiPage cache indices so their structural boundaries and container ownership
can be validated.

An unsupported class remains an opaque control only when its site record and
`ObjectStreamSize` unambiguously delimit all of its bytes. Unknown fields
inside a known record are accepted only when its declared record boundary still
reconciles exactly; otherwise parsing fails.

The serializer is deliberately no-op and lossless. It replays the
retained `f`, `o`, optional `x`, `\x01CompObj`, `\x03VBFrame`, opaque streams,
and nested storage metadata. It then reparses that subtree with the supplied
project code page before returning it. The decoded and raw model is bound to
its read-time persistence signature; changing either side returns
`ErrUnsupportedMutation` rather than silently discarding an edit. Controlled
property compilation uses the explicit mutation boundary below; new Designer
generation remains a separate stage.

`vbaproject.Project.Forms` owns parsed Designer subtrees. Their streams and
storage metadata are excluded from generic `RawStreams` and
`StorageMetadata`, then restored by `vbaproject.Write` through `SerializeForm`.
Unrelated root streams remain opaque pass-through data. Projection is owned by the separate
`internal/vba/userforms/projection` adapter so the lossless model does not
depend on the user-facing schema.

## Encoding and geometry

The string storage contract follows [MS-OFORMS Strings](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms/83c93081-0417-4f51-8e2d-9247af5003fa)
and [String Compression](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms/b9d3da32-3d53-4f98-b0cb-c64af7f13457).

Compressed MS-OFORMS strings store the low byte of UTF-16 code units whose
high byte is zero; they do not use the VBA project code page. Japanese and
other characters outside U+0000..U+00FF require uncompressed little-endian
UTF-16. An odd uncompressed byte count fails. `\x03VBFrame` uses the VBA
project code page; unsupported code pages and undecodable bytes fail. CompObj identity
strings use the same supplied project code page while their complete bytes are
also retained.

Control positions and sizes remain signed HIMETRIC integers in the binary
model. Conversion to FormSpec points belongs to the projection layer.

## FormSpec projection

`projection.Project` converts one parsed `oforms.Form` to the canonical
`internal/vba/userforms/spec.FormSpec`. The result uses schema version 1,
`kind: xlflow.userform`, `basis: designer`, and parent-relative point
coordinates. HIMETRIC geometry is converted with `points = value * 72 / 2540`.

Controls are flattened in persisted preorder. IDs use deterministic
`control_NNN` values, `parentId` points at the containing control, and `zIndex`
is the persisted sibling index. Known MSForms classes receive their canonical
type and ProgID. Caption, text/value, size, position, tab index, ComboBox/ListBox
selected index, enabled state, and visible state are projected when their binary
meaning is known. File-format defaults are applied when the relevant MS-OFORMS property
record is omitted.

The root `DisplayedSize` is exposed through `form.build.clientWidth` /
`clientHeight` and the matching observed client fields. It is the persisted
client/display area and can differ from Excel VBIDE's outer Designer width/height
because of window chrome. Projection does not label it as outer dimensions.
Control geometry is compared directly. Boolean control values are normalized from persisted
`1`/`0` spellings to Excel snapshot `True`/`False` spellings.

TabStrip is structurally recognized by the binary reader, but its persisted
`ListIndex` is not part of the current FormSpec control contract. Projection
therefore records it as unsupported rather than emitting a snapshot the
canonical loader would reject.

Binary masks, padding, TextProps, class tables, opaque tails, stream extents,
and internal site IDs never enter FormSpec. Semantically meaningful state that
has no supported FormSpec field is listed in a control's sorted `unsupported`
array and summarized by one `unsupported_properties` warning per control.
Non-empty site strings without a FormSpec field—including `Tag`,
`ControlTipText`, `RuntimeLicKey`, `ControlSource`, and `RowSource`—are named in
that list rather than silently discarded. Site `Name` is already represented
by the control's canonical `name` field.
Unmodeled `BooleanProperties`, control flag bits outside the projected
`Enabled` bit, and site flags that differ from the persisted default are also
reported. For nested containers, their level-specific font, picture, mouse icon,
and extra streams are attached to the owning control's unsupported list and
warning.
Form-level unsupported state is summarized by a form warning. Raw persistence
details used only for lossless replay are not exposed as property-bag values.
The file-pull publisher additionally appends the operational warning
`compatibility_artifact_unsynchronized` to every emitted FormSpec. This marker
does not describe a projection loss: it records that pure-Go pull deliberately
did not regenerate compatibility `.frm` / `.frx` artifacts. Shared source
preflight rejects the marker with `FRM201` before either push backend can import
a missing or stale `.frm`; `form build` followed by an Excel-backed pull clears
the unsafe state by applying the spec and refreshing those artifacts.

A structurally bounded control whose type or ProgID cannot be recovered is
retained as `type: Control` with `unsupported: [controlType]`. `UFV015` accepts
that shape only as a snapshot-only placeholder. `form build` and `form apply`
reject it with `UFV006` before opening Excel; callers must replace it with a
supported type or the real custom ProgID before authoring.

## Controlled FormSpec mutation (Issue #882)

`internal/vba/userforms/compiler.CompileEdits(base, before, after, codePage)`
accepts two complete snapshots of one existing form. It applies only changed
authoring fields and returns an independently owned, reparsed `oforms.Form`
that the existing `SerializeForm` and `vbaproject.Write` can publish. It does
not consume `.frx`, open Excel, or enable Designer edits in the `pack` CLI;
CLI integration remains Issue #887.

Control IDs connect the before/after snapshots. Names and the persisted
parent/sibling relationships connect the before snapshot to the binary model.
Control creation, deletion, rename, type/ProgID changes, reparenting, and
reordering are rejected. Snapshot-only `observed`, `warnings`, and `unsupported`
metadata is not mutation input. The before value of every changed authoring
field must agree with the binary model when that value is supplied; a stale
baseline fails instead of overwriting a different persisted value.
Geometry comparisons allow the existing 0.05-point Designer parity tolerance
for Excel twip versus persisted HIMETRIC rounding.
Both snapshots must use an omitted coordinateSystem, `points`, or
`parent-relative`; all three describe point-valued geometry relative to the
owning parent. Other coordinate systems are invalid even for a no-op and are
rejected before numeric conversion, not silently interpreted as points.

The initial editable classes are Label, TextBox, ComboBox, ListBox,
CommandButton, CheckBox, OptionButton, and existing Frame containers. Their
applicable caption, text/value, parent-relative position, size, tab order,
enabled, and visible fields use explicit persistence mappings. Existing nested
control geometry remains parent-relative. Form caption is supported; root
width/height edits are rejected because persisted client dimensions and Excel
outer Designer dimensions are different. MultiPage/Page/TabStrip editing and
container topology generation remain subsequent stages.

For the Excel-authored unbound ComboBox/ListBox fixtures, Designer AddItem
items and ListIndex disappear after save/close/reopen. ComboBox Value persists,
but ListBox Value returns to null without persisted list state. Therefore
list/selectedIndex edits and ListBox text/value edits are rejected. RowSource
and runtime list initialization are outside this compiler's supported layouts.
The observation is retained under the compiler's Excel-authored testdata.

TextBox text and value address the same persisted Value string. A change to
one alias is sufficient; conflicting changes to both aliases fail. The same
rule applies to top-level control fields and their property-bag aliases.
An explicitly changed `form.build.caption` takes precedence over the legacy
form caption. Unchanged build/observed fields do not override a changed field.
When build.caption first appears, the supplied legacy before caption remains
the stale-input baseline; an existing explicit before build.caption takes
precedence over the legacy baseline. Observed metadata is never synthesized
into authoring intent or a supplied baseline.
Empty strings, zero, and false are explicit values. Removing a property to
request a default reset is unsupported.

The supported property bag is case-insensitive and rejects duplicate aliases.
It includes applicable common-field aliases plus Tag, ControlTipText,
GroupName (OptionButton), BackColor, ForeColor, BorderColor, BorderStyle, and
MaxLength (TextBox/ComboBox). A binary property must actually exist in the
control's persistence table. Unsupported property bags remain opaque when
unchanged; changing an unsupported property is an error. Colors use unsigned
32-bit OLE_COLOR values, BorderStyle is 0 or 1, and MaxLength is a non-negative
signed-32-bit value. Geometry uses nearest-integer HIMETRIC conversion; width
and height are non-negative and all geometry must fit signed 32-bit storage.
TabIndex is 0..32767. CheckBox/OptionButton checked values accept Boolean,
True/False, or 0/1/-1 spellings; tri-state null edits are not supported.

Compilation errors include a stable code, form, control, property path, and
reason. Codes distinguish `userform_edit_unsupported`,
`userform_edit_invalid`, `userform_edit_stale_input`, and
`userform_edit_conflict`. Any rejected edit fails the entire compilation.
No input model or snapshot is changed, and no partial output model is returned.

The low-level `oforms.ApplyEdits` boundary validates the original read-time
signature, owns a private clone, and encodes only changed records/sites.
Masks, string byte lengths, alignment, record lengths, ObjectStreamSize, and
site-data CountOfBytes are recalculated. Untouched records, TextProps, images,
opaque tails, class/depth tables, auxiliary streams, stream-name spelling, and
storage metadata are retained. MS-OFORMS strings keep their existing encoding
when it can represent the new text losslessly and can use UTF-16 otherwise.
Unknown layouts whose mutation cannot be justified are rejected. Reparse and
edited-value comparison complete before a signed model is returned; the
ordinary serializer never blesses arbitrary caller changes.

Enabled uses VariousPropertyBits bit `0x2` for embedded controls and
BooleanProperties bit `0x4` for Frame. Omitted VariousPropertyBits uses the
MS-OFORMS class-specific default, independently of flags Excel explicitly
stores for a new control. Every other flag bit is preserved. The committed
baseline/all-disabled fixtures bind these changes to saved/reopened Excel
Designer state, including nested TextBox and Frame.

## New Designer generation (Issue #883)

`compiler.CompileNew(spec, codePage)` authors a new flat Designer without a
template, `.frm`, `.frx`, Excel, COM, or VBIDE. `oforms.NewForm` accepts a
separate persistence-unit definition and reuses the existing property tables,
record encoders and CFB writer. It emits required `f`, `o`, `\x01CompObj`, and
`\x03VBFrame` streams, then reparses them before returning a signed model.
No arbitrary caller-created model is accepted by the lossless serializer.

The supported classes are Label, TextBox, CommandButton, CheckBox,
OptionButton, ToggleButton, ComboBox, ListBox, SpinButton, ScrollBar, and
Image without Picture data. Frame and all nested/container structures,
custom ActiveX, resources, list/selectedIndex state, and unimplemented
properties fail loudly. ListBox text/value is rejected because it depends on
unpersisted list state. SpinButton/ScrollBar value is an integer in the
initial supported range 0..100; their binary Position is used instead of a
MorphData Value string. Boolean controls include ToggleButton and accept the
same two-state spellings as the existing compiler.

Geometry uses points converted to the nearest signed-int32 HIMETRIC value.
Root `form.build.clientWidth` / `clientHeight` address client dimensions,
with defaults of 240 by 180 points. Outer width/height input is unsupported,
and supplying both dimension families is invalid. Snapshot observations
alone are not root authoring intent.

These fields mean the persisted MS-OFORMS `DisplayedSize` / VBFrame
`ClientWidth` and `ClientHeight`, not the runtime MSForms `InsideWidth` and
`InsideHeight` properties. The focused Excel gate observed a 4.55-point
height difference on both Excel-authored and generated forms. No constant
offset or outer-to-client estimate is applied. The gate compares exported
VBFrame client dimensions and reads the normalized binary after saving.

The default control sizes are:

| Classes                                     | Width x height (points) |
| ------------------------------------------- | ----------------------- |
| Label, CheckBox, OptionButton, ToggleButton | 72 x 18                 |
| TextBox, ComboBox                           | 120 x 18                |
| CommandButton                               | 72 x 24                 |
| ListBox                                     | 120 x 72                |
| SpinButton                                  | 18 x 36                 |
| ScrollBar                                   | 120 x 18                |
| Image                                       | 72 x 72                 |

Position defaults to zero. Controls are ordered by normalized zIndex with
stable input-order ties. Site IDs start at 1 in that order; NextAvailableID
is one past the maximum (1 for an empty form). Default tab order follows
the same order and explicit indices are retained. Site counts, byte counts,
record lengths, and object extents are computed from encoded bytes.
MS-OFORMS captions and values use compressed low-byte UTF-16 or full UTF-16,
independent of the project code page. VBFrame remains code-page text and
does not duplicate Unicode captions from the authoritative FormControl.

Errors retain form/control/property context and use
`userform_generation_invalid`, `userform_generation_unsupported`, or
`userform_generation_conflict`. Rejected input returns no partial Designer.
Caller input is never mutated.

New code-behind is generated from code-only text with form attributes and
two new component GUIDs in VB_Base; class identities are fixed constants.
Project addition ensures a Microsoft Forms reference and preserves existing
reference groups verbatim. Missing references use a canonical REGISTERED
Forms 2.0 LIBID with fixed `C:\Windows\System32\FM20.DLL` path, independently
of the build host. Declaration,
module and PROJECTwm bookkeeping are completed by the project writer;
blank pack uses canonical specs/code; template Designer integration remains #887.

REGISTERED reference identity must come from a canonical Forms LIBID in a
complete sized record. CONTROL reference identity must come from the original
TypeLib GUID at its declared position in the complete extended record. A
REFERENCEORIGINAL, display name, or twiddled/extended LIBID alone is insufficient,
including a Forms LIBID paired with a foreign OriginalTypeLib or a GUID appearing
only in a file path or description. Project addition independently rejects
equal component GUIDs in `VB_Base`, comparing them case-insensitively even
when the caller constructs the module without `NewUserFormModule`.

## Structural validation

The reader and serializer return `oforms.ErrMalformed`, wrapped by an `oforms.ParseError`
carrying storage path, stream, byte offset, and structure, when any of these
invariants fail:

- required top-level `f`, `o`, `\x03VBFrame`, or `\x01CompObj` streams are
  missing;
- a version, length-prefixed field, record boundary, or alignment step leaves
  its enclosing stream;
- expanded `SiteDepthsAndTypes` entries do not equal `CountOfSites`;
- `CountOfBytes` does not end at the expected site-data boundary;
- the sum of site `ObjectStreamSize` values differs from the `o` stream size;
- a container site lacks its matching `i<site-id>` storage, a non-container
  claims one, two storages claim the same ID, or a child storage has no owner;
- a known property record leaves unexplained bytes outside an explicitly
  supported opaque tail.

The resource limits are 64 MiB per designer stream, 65,535 sites per form, and
64 nested container levels. Counts and lengths are checked against remaining
input before allocating or slicing. Serialization validates all forms before
the enclosing CFB is published. A malformed retained count, length,
`ObjectStreamSize`, or nested ownership relationship therefore fails instead
of producing a partial `vbaProject.bin`.

## Verification

The committed Excel-authored `p4_form.bin` fixture covers a simple form and
known control property record. `p6_nested_form.bin` covers Frame, MultiPage,
Page, and nested container storages. Focused corruption tests cover missing and
mis-sized `o` streams, inconsistent depth runs, orphan storages, unsupported
code pages, and structurally bounded opaque controls. Native Go fuzz targets
exercise the public CFB-to-form reader, successful parse-to-serialize replay,
and the `f` stream parser. Serializer tests compare every simple and nested
Designer stream and storage metadata entry byte-for-byte, cover CP932 text,
and reject malformed retained data and unsupported semantic mutation.
Projection tests cover common properties, nested parent relationships,
snapshot-only opaque controls, validation boundaries, and repeated byte-stable
JSON output. `TestProjectMatchesExcelBackedSnapshot` can additionally compare
the canonical supported control state from an Excel-authored snapshot with the
pure-Go projection of the same workbook. Set
`XLFLOW_PROJECTION_PARITY_WORKBOOK` and
`XLFLOW_PROJECTION_PARITY_SNAPSHOT` to run this Windows/Excel integration test.
Root form dimensions are excluded because the binary stores client dimensions
while Excel reports outer Designer dimensions.

Compiler tests cover edits, string growth/shrink, untouched controls and
subtrees, stale input, alias conflicts, range errors, and atomic failure.
`scripts/test-userform-mutation-e2e.ps1` is a developer-only Excel gate:
its create phase records persistence observations; its verify phase compares
compiler output against expected Designer properties and runs a sentinel.
Verify executes workbook VBA with the developer's Excel authority: use only
the trusted generated baseline and known compiler-produced derivatives.
Read-only opening and disabled events do not sandbox the explicit sentinel;
property comparisons do not authenticate workbook code. Do not run verify on
an untrusted workbook in a credential-bearing developer environment.
An optional new `NormalizedWorkbookPath` saves and reopens an independent
artifact in the same owned Excel instance. The gated Go tests
`TestGenerateExcelMutationArtifact` and `TestReadExcelNormalizedArtifact`
prepare the file-level output and validate the normalized VBA project. They
are skipped unless their documented `XLFLOW_MUTATION_*` environment variables
are supplied. These Excel gates are never ordinary CI tests.
