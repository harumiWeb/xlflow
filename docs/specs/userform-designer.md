# VS Code UserForm Designer

This specification defines the interactive UserForm Designer in the
xlflow Visual Studio Code extension. The open VS Code `TextDocument` remains
the project source. The Designer keeps pending source edits in one
URI-shared `CustomDocument` draft and previews that draft through the xlflow
Language Server Protocol (LSP). The draft preserves the source text and is
saved through the VS Code document lifecycle.

## Editor selection and document eligibility

The extension registers `xlflow.userFormDesigner` through
`CustomEditorProvider` as an optional custom editor for `.yaml`, `.yml`, and
`.json` files. Its `option` priority keeps the normal text editor available.
Users can choose the visual editor with **Reopen Editor With...**.
Opening the file with the Designer redirects to its dedicated
`xlflow-userform-designer` editor resource. The resource maps back to the
original file URI, which remains the LSP and source-edit target. Separate
resource URIs isolate Designer native Undo from the regular text editor's
history and save-time source edits.

The preview is available only for a FormSpec directly inside the `specs`
directory under the configured forms root (`[src].forms` in `xlflow.toml`,
defaulting to `src/forms`). Arbitrary YAML, YAML nested in another directory,
and other JSON files are not UserForm documents. The file must belong to an
xlflow project with `xlflow.toml` and be handled by the currently selected
project's single xlflow LSP client. Opening a document associated with another
project, or without a compatible LSP connection, produces an editor error.

## Deferred-save draft and document lifecycle

The provider owns one custom document and source-text draft per eligible
FormSpec URI. Every Designer panel for that URI reads and edits the same draft
and native history. The draft starts from the latest open `TextDocument` text
and preserves comments, formatting, and other untouched source. Persist pending
drafts in extension internal storage keyed by the original FormSpec URI so they
survive normal close, reopen, and extension restart.

The native `CustomEditorProvider` save callback durably checkpoints the draft
to extension internal storage only. AutoSave, File Save All, and menu Save use
this callback and never write YAML, YML, or JSON. The native tab dirty marker
therefore tracks the internal checkpoint, while a separate **YAML/JSON not
applied** badge indicates that the draft has not been published to source.
Neither the extension nor the Designer changes the user's global
`files.autoSave` setting. Regular source-text editing and its configured
AutoSave behavior continue normally; edits in that document appear in the
Designer immediately.

Only the Designer Save button and its Designer-scoped Ctrl+S/Cmd+S command
publish the current validated draft. Reconcile against the newest open
`TextDocument`, apply source-preserving edits through VS Code document APIs,
and call `TextDocument.save()`; never write the workspace file directly. On
successful publication, refresh the internal checkpoint and clear pending
recovery state. A failed validation, conflict, edit, or source save retains the
pending draft for retry.

The host sends the exact current draft text and its revision to the LSP for
preview, canonical validation, and semantic edits. It applies the localized
text edits returned by the LSP to the draft. The LSP and canonical FormSpec
parser remain the only authority for FormSpec meaning and validation; neither
the host nor Webview adds independent property or schema rules.

Changes made in the regular text editor are observed immediately and reflected
in the Designer. If the Designer draft has no pending edits, it adopts the
current text. If it is dirty, the draft engine merges non-overlapping changes
against the last synchronized source and rebases draft history. An overlapping
change retains both the current text-editor source and the Designer draft,
shows a source conflict, and disables explicit source publication until
resolved. Native checkpoint saves remain available. The
user can choose **Use YAML/JSON changes** or **Use Designer draft**. When a
safe granular merge is unavailable, choosing a whole-document version
requires explicit confirmation. Neither version is silently discarded.
Conflict detection conservatively treats different edits on the same physical
source line as overlapping, including compact JSON and YAML flow mappings.
Changes on different lines are merged when their source ranges do not overlap.
Before merging a dirty draft, the host compares canonical LSP projections of
the baseline and latest source. Changes to control order, ID, type, ProgID or
parent require explicit conflict resolution so geometry cannot move to another
control after a source reorder. Invalid source or an unavailable canonical
projection also retains both versions rather than assuming a merge is safe.
Conflict resolution rebases the existing draft history onto the latest source
and adds one native Undo step. Undoing both the resolution and the earlier
Designer action returns to that latest source baseline. History entries
superseded by later direct source edits are neutral native callbacks; advancing
past them never restores discarded source or throws a stale-frame error.

Explicit Designer Save first validates and reconciles against the latest open
`TextDocument`, applies source edits through VS Code document APIs, then calls
`TextDocument.save()`. A stale draft never overwrites a newer dirty text-editor
buffer, and the provider does not overwrite workspace source through filesystem
writes. Save As is an explicit publish to the chosen destination; it creates a
missing destination through VS Code, and an existing dirty destination must be
saved or reverted first. Backup, restore, hot-exit, and Revert use the
custom-document lifecycle. Revert discards the pending Designer draft and
adopts the latest source `TextDocument` buffer, including unsaved text-editor
changes; clear recovery state only after that succeeds. Advancing past
superseded neutral Undo entries can conservatively dirty the custom document
and trigger another internal checkpoint. The native tab dot tracks that
checkpoint, while the source-not-applied badge remains until publication.
AutoSave follows the standard VS Code configuration and checkpoints internal
storage only.

A corrupt checkpoint displays a recovery error and remains stored; it is never
silently discarded or overwritten. Clear recovery state only after explicit
source publication or Revert completes successfully.

Each completed Designer action raises one `CustomDocumentEditEvent`.
VS Code Undo/Redo updates the shared draft, and the draft engine rebases its
history when external text changes are merged.
VS Code backup and hot-exit recovery remain supported in addition to the
URI-keyed internal checkpoint. Restoring a checkpoint recovers the pending
source draft; source publication remains an explicit Designer action.

## Source and preview protocol

The extension host reads the URI-shared custom-document draft, including
merged unsaved text-editor changes. It sends the document URI, current draft
revision, and full text to the LSP. The Webview never reads the document or
workspace files directly.

The server advertises preview support through
`capabilities.experimental.userFormPreview: true`. The request is:

```json
{
  "uri": "file:///project/src/forms/specs/Main.yaml",
  "version": 12,
  "text": "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: Main\ncontrols: []\n"
}
```

The `xlflow/userFormPreview` response always includes the requested `version`
and may include a parsed canonical `document`, validation `warnings`, or an
`error`:

```json
{
  "version": 12,
  "document": { "form": { "name": "Main" }, "controls": [] },
  "warnings": []
}
```

An error has a stable `code` and a human-readable `message`; `line` and `column`
are included when the parser can identify a source location. The language
server validates the URI and FormSpec kind and parses the supplied text with
xlflow's canonical FormSpec parser. It does not replace the supplied text with
saved file contents or invoke Excel.

The extension refreshes the preview when the document changes or the LSP
connection becomes available again. Results whose document version or
connection generation is no longer current are discarded. Warnings are shown
alongside a valid preview. If parsing, validation, project selection, or the
LSP request fails, the Webview displays the error and retains the last valid
rendered document. A later valid document automatically replaces that state.

The host-to-Webview protocol carries a rendered `document`, an
`invalidDocument` status, or a `themeChanged` notification. The Webview can
signal that it is ready and can ask VS Code to open the backing document in the
normal text editor. Completed geometry interactions send semantic edits; the
Webview never serializes or replaces the source document.

## Rendering

The Webview uses Preact and TypeScript to render the FormSpec form and its
controls. It displays the built-in controls supported by the FormSpec model,
including Label, TextBox, CommandButton, CheckBox, OptionButton, ComboBox,
ListBox, ToggleButton, SpinButton, ScrollBar, Image, Frame, MultiPage, Page,
and TabStrip. Unknown control types receive a textual placeholder. Control
positions and sizes are an approximate visual representation, not a pixel
accurate reproduction of the VBA designer.

Dimensions are represented in points and displayed at 96/72 CSS pixels per
point. Missing control positions default to zero. Missing control sizes use
these display defaults, in points:

| Control                                     | Width × height |
| ------------------------------------------- | -------------: |
| Label, CheckBox, OptionButton, ToggleButton |        72 × 18 |
| TextBox, ComboBox                           |       120 × 18 |
| CommandButton                               |        72 × 24 |
| ListBox                                     |       120 × 72 |
| SpinButton                                  |        18 × 36 |
| ScrollBar                                   |       120 × 18 |
| Image                                       |        72 × 72 |
| Frame                                       |      144 × 108 |
| MultiPage                                   |      240 × 180 |
| TabStrip                                    |       240 × 48 |
| Unknown control                             |        72 × 24 |

If the form has no usable dimensions, its display size defaults to 240 × 180
points. The preview marks dimensions supplied by these defaults as approximate.
Page content is laid out inside its parent MultiPage; the Page is not rendered
as an independently positioned top-level control. Visible Page headers switch
the active preview Page. A Page list in the property region exposes Pages by
stable control ID, including Pages whose headers are hidden. Selecting a header
or list entry updates only transient `activePageId`; it does not edit source or
write `selectedIndex`. The preview initially follows authored selection, or
defaults to the first Page when no selection is recorded.

Explicitly invisible controls and their descendants are not rendered. Invisible
Page headers are omitted, but the Page list can still select those Pages for
inspection and property editing. Selection indexes retain the full ordered
Page collection. A selected invisible Page does not fall through to a
different Page. Omitted visibility defaults to visible.

Page collection order follows stable sibling `zIndex` order, matching the
pure-Go generation and topology compiler. Selection is applied before hiding
invisible Pages. The preview reads root `form.width` and `form.height` directly;
it does not infer authoring dimensions from observed client, inside, or outer
Excel values. An omitted axis uses the preview default and remains approximate.
ComboBox display prefers
explicit text, then value (including empty strings), then the first list item.

The extension host supplies localized UI strings through a `localization`
message when the Webview reports readiness. These are rendered as text; parser
diagnostics and workspace-authored captions retain their original wording.

Page tabs can select a Page for property inspection. The Designer does not
load image assets. Image controls use a placeholder even when a
picture path is present in the FormSpec.

Source-preserving semantic edits use the host-neutral Go API. See
[`userform-semantic-edit.md`](userform-semantic-edit.md) and
[`ADR-0065`](../adr/ADR-0065-userform-source-preserving-semantic-edits.md).

## Interactive geometry

Excel-backed `form build` and apply verification allow the native MSForms
rounding of control `Left`, `Top`, `Width`, and `Height`: half a twip
(0.025 pt) plus one single-precision ULP at the requested value. Excel 16.0
returns these fields as `Single` after rounding to twips (0.05 pt).
This tolerance applies only to the four geometry fields; numeric `Value`,
`TabIndex`, Boolean and string persistence checks retain their existing
strictness. Non-finite geometry and differences outside the bound fail with
`designer_write_failed`. This does not round or rewrite authored FormSpec
values and does not alter pure-Go HIMETRIC conversion. Regression coverage
includes fractional Designer values, half-twip boundaries, larger coordinates,
ignored writes and strict numeric-value checks.

Selection is a transient control ID; clicking the background selects the form.
Known built-in controls other than Page can move and resize with eight handles.
When the form is selected, its right edge, bottom edge and bottom-right corner
resize the persisted client width/height. Window decorations are preview-only
and are never added to authored dimensions. Form resizing requires property
editing capability and numeric width/height descriptors from LSP metadata.
It preserves child coordinates and protects all immediate root child extents,
including hidden controls. Page geometry is not edited on the canvas. Unknown controls remain
selectable placeholders. Only a completed parent-changing pointer drag can
commit a new relationship; overlap alone never changes source.

Pointerdown captures the original parent-relative geometry. Pointermove only
updates local preview state; pointerup emits at most one transaction. North/west
resize can emit both `moveControl` and `resizeControl` in that transaction.
Escape, pointer cancellation, capture loss, pointer focus loss and document
updates cancel unfinished interaction. Pointer cancellation and capture loss
only cancel the matching active pointer; other pointer IDs cannot discard its
preview or pending transaction. Ignore additional pointerdown events while a
pointer gesture is active so that selection and gesture ownership remain with
the original pointer. Arrows move 1 pt, Shift+Arrow moves
10 pt; repeats accumulate until all movement keys are released or focus leaves
the canvas. No-op interactions create no document edit.

Use `96/72` CSS pixels per point and a common inverse transform for pointer
deltas. Zoom defaults to 100% and offers 50/75/100/125/150/200%. Container
content insets are point-based and shared by rendering and constraints:
Frame uses 0.75 pt per edge; MultiPage uses 1.5 pt left/right/bottom and
18 pt for its tab band. These approximate UI insets are not authored geometry.
Selection, zoom, grid visibility and snapping are not persisted into FormSpec.

The 8 pt grid is initially visible; snapping is initially off. Snap dragged
coordinates/resize edges before applying parent boundary constraints, which
take precedence over the grid. Keyboard steps ignore snapping. Compute geometry
from initial values and total displacement, with displacement rounded to a
millionth of a point, rather than repeatedly converting CSS positions.

Completed interactions must fit inside the parent's displayed content area,
with a 1 pt minimum on both axes. Frame shrink protects immediate children;
MultiPage shrink protects children on every Page, including hidden/non-selected
Pages. Empty containers, including MultiPage controls with empty Pages, retain
the 1 pt outer minimum; child extents plus decoration insets increase that
minimum only where children require space. Moving a container retains child local coordinates. Existing overflow
is never fixed on load. Container children are visually clipped to the parent's
content bounds so visible pointer targets agree with drop hit testing; source
coordinates remain unchanged. Fitting moves/resizes are permitted, but an oversized
control cannot move until resized to fit. A control smaller than 1 pt on either
axis also cannot move until explicitly resized to meet the minimum; source
dimensions are never silently enlarged by moving it. If child protection prevents fitting,
editing that geometry is unavailable; source editing remains available. These
UI constraints do not add canonical FormSpec validation rules.

Move and resize preserve axes with zero displacement, including when snapping
is enabled; a zero-displacement gesture produces no edit or history entry.
Resize also preserves the opposite edge. If that unchanged
axis or opposite edge already prevents a fitting result, reject the interaction
rather than silently correcting it; source editing remains available.

## Multi-selection and productivity commands

Group selection is limited to ordinary built-in controls that are immediate
siblings under one parent: the form root, one Frame, or one Page. Ctrl/Cmd-click
and Shift-click toggle eligible siblings in or out of the group. Custom
controls remain selectable alone but cannot join a group. A marquee selects
only visible, ordinary built-in immediate children whose entire displayed
bounds are inside the marquee. A group cannot span parents. Show a group
boundary for multi-selection; keep the last selected control as the explicit
anchor. Changing to a different parent starts a new selection.

Dragging a multi-selection applies the same parent-relative translation to
every member, preserving the parent and all relative geometry. Multi-selection
drag has no resize handles and never resizes its members. The move stays within
the current parent's content bounds; apply any boundary correction to the
whole group so relative geometry remains fixed. Reject a group that cannot fit.

Provide Align Left, Center, Right, Top, Middle, and Bottom. Each aligns selected
siblings to the corresponding edge or center of the last selected anchor.
Provide Make Same Width, Make Same Height, and Make Same Size using the anchor's
current dimensions. Provide Distribute Horizontally and Distribute Vertically
for at least three siblings: order by left or top coordinate, using stable
z-order to break ties; retain the first and last controls' positions on the
chosen axis and divide the intervening edge-to-edge gaps equally. Provide
Bring to Front, Bring Forward, Send Backward, and Send to Back. These z-order
commands preserve the selected controls' relative order and do not alter
`tabIndex`. Layout operations are atomic and must reject a result that cannot
meet the existing parent-bound constraints without changing its requested
alignment, size, or distribution.

Expose Cut, Copy, Paste, Delete, Duplicate, z-order commands, and Properties in
the selection context menu. Copy and Cut request source records through the
LSP. A successful Copy replaces the shared in-memory `ClipboardSnapshot`; Cut
keeps its captured snapshot provisional until deletion succeeds. Paste and
Duplicate allocate new IDs and names and create one source edit. A copied Frame
includes descendants. A copied MultiPage includes its
Page controls and their descendants; a Page cannot be copied directly. Reject
an unsupported custom-control subtree as a whole. Preserve source-present
fields and observations exactly, distinguishing absent, null, empty string,
false, and zero; do not manufacture preview or normalization defaults. TabStrip
`tabs` stay attached to their TabStrip control. Image bytes, event handlers,
and code-behind are never copied. Same-project authored picture paths remain
unchanged; reject a cross-project paste when the copied subtree contains one.

Offset each repeated paste by 8 pt on both axes. Translate the complete copied
group uniformly to fit the destination's content bounds; reject it if the full
group cannot fit. Do not resize pasted controls. Paste uses a selected Frame or
Page as its destination only when exactly one control is selected; otherwise it
uses the selection's owning parent. Duplicate always uses the original
selection's owning parent, so a Frame is not duplicated inside itself.

Cut captures a valid source snapshot first, then asks one confirmation with the
aggregate descendant count for all selected roots in the same per-URI edit
queue. Leaf-only Cut needs no descendant confirmation. If confirmed and still
current, Cut submits one atomic `removeControls` cascade edit. Only after that
edit succeeds does the host publish the captured snapshot to the shared
clipboard.
Cancellation, stale source, failed capture, or failed deletion leaves both the
source and previous clipboard unchanged. A request superseded before deletion
is rejected without deleting its source.

Each command is one productivity semantic operation and one native VS Code
`CustomDocumentEditEvent`/Undo step. Copy extraction is read-only. Cut captures
its snapshot before confirmation, keeps that snapshot private during confirmation,
and publishes it only after the single cascade edit succeeds. Productivity
operations are not combined with each other or legacy operations in one
`userFormEdit` transaction.

## Toolbox and structural control edits

The Toolbox contains Pointer and these 14 built-in control types: Label,
TextBox, ComboBox, ListBox, CommandButton, CheckBox, OptionButton,
ToggleButton, SpinButton, ScrollBar, Image, Frame, MultiPage, and TabStrip.
Page is not a Toolbox item. Pointer is the normal selection tool. Selecting a
control type arms placement; clicking the canvas inserts one control at the
click position. Drag-to-create is not part of this interaction.

With `userFormParentEdit`, placement targets the innermost visible Frame under
the pointer, or the form background. When `userFormPageEdit` is also
advertised, insertion on the active Page targets that Page. Without the needed
parent/Page capabilities, new controls remain root controls. Frame and Page
placement explicitly assign `parentId`; overlap does not change existing
relationships. Use the same centralized per-type width and height defaults listed under Rendering;
creation and preview must not maintain separate size tables. Convert the
canvas click to FormSpec points. With snapping off, normalize each coordinate
with the TypeScript `stableDelta` helper (six decimal places) to suppress
floating-point micro-error; with snapping on, use the existing 8 pt grid. Then
clamp `left` and `top` so the default-size control fits within the current form
bounds or the target Frame/Page content bounds. If either destination dimension is smaller than the new control's default
dimension, reject the insertion without changing the source.

Generate the ID as `control-<UUID>`, independently of its current name, and
ensure it does not collide with an existing ID. Renaming a control leaves this
ID unchanged. Generate a familiar VBA name from the control type's prefix by
choosing the smallest positive integer suffix whose full name is unused when
compared case-insensitively against every control name and the form name. Before
dispatching an addition, the extension host rechecks the new ID and checks the
name case-insensitively against the current form and control names. These
Toolbox allocation and host-add checks are the Designer's name-conflict guards;
Go `Apply` and canonical FormSpec validation do not add global control-name
uniqueness validation.

A new MultiPage has no Page child controls and `selectedIndex: -1`. A new
TabStrip has `tabs: []` and `selectedIndex: -1`. Inserting either control does
not create or edit Pages or tabs. On successful insertion, select the new
control and return the Toolbox to Pointer. Escape cancels pending placement
without a source edit.

The Designer offers a Delete toolbar action and handles Delete when the canvas
has focus. It must not intercept Delete from a focused text or property input.
The keyboard event must target the canvas itself: focus on a Page-tab button
must not delete a previously selected control through event bubbling.
A control with no descendants is removed immediately. If descendants exist,
the host first shows a modal confirmation containing the selected control's
name and descendant count. Confirmed removal deletes the selected control and
all descendants with `cascade: true` in one semantic edit and one VS Code
`CustomDocumentEditEvent`; Undo/Redo therefore restores or removes the whole
subtree as one native history step. The canvas Delete action does not delete a Page.
Page deletion is available from the Page list and always requires explicit
confirmation identifying the Page and the number of descendant controls; the
confirmed action removes that Page and its descendants with `cascade: true` in
one semantic edit and one native history step.

Structural edits apply only to the shared FormSpec draft. They do not edit
code-behind or publish the source document until explicit Designer Save or its
scoped keyboard command. The
host captures the draft source and revision associated with an add request or
pending delete confirmation. If that draft revision, project, or LSP
connection changes before the action is submitted or applied, discard the
stale action and refresh from the current merged draft; do not replay the
request against newer text. Already committed draft history follows the
external-change rebase policy above.

## Frame parenting and hierarchy

The LSP advertises `capabilities.experimental.userFormParentEdit: true`.
The extension forwards it as `parentEditable`. Parent-changing pointer drags
require both parent and geometry capability; additions with `parentId` require
both parent and structural capability. Missing capability retains the existing
same-parent movement and root-only placement.

Hit testing follows rendering's stable sibling z-order and visible hierarchy.
The innermost Frame content region beneath the pointer is the destination;
ordinary children inherit their Frame destination. With
`capabilities.experimental.userFormPageEdit`, an active Page is also a valid
destination for ordinary controls. The moving control and its descendants are
excluded. A Page can be chosen as the destination by inserting on the active
Page, dropping into its content, or using the Page destination picker. A Page
control itself cannot be reparented or moved to another MultiPage. An existing
Frame inside a Page can still be an explicit Frame destination. The form
background permits root parenting; invalid destinations never fall back to
root.

Frame destinations require an omitted/empty ProgID or a case-insensitive match
to `Forms.Frame.1`. Custom-ProgID Frames remain visible but cannot themselves be
placement or reparenting destinations; preview targeting, creation and host
preflight share this predicate, with LSP validation remaining authoritative.

World-space origin conversion includes every ancestor's local position and
content insets. Page origins derive from MultiPage content. The drag preview
uses a form-coordinate overlay, preserving the pointer offset and all child
local coordinates. On changing parents, convert to destination-local points,
apply grid snapping and six-decimal normalization, then clamp to fit. Oversized
or sub-1-point controls cannot be dropped. Invalid drops cancel without edits;
unchanged clicks do not reparent. Resize and keyboard movement retain their
current parent and existing constraints.

A completed parent change sends exactly one semantic operation:

```json
{ "type": "setParent", "controlId": "label", "parentId": "frame", "left": 12, "top": 18 }
```

All five fields are required by the Designer RPC; `parentId: ""` means the
form root. Same-parent movement uses `moveControl`. Frame movement updates only
that Frame, while children retain their local geometry. LSP and host reject
Page subjects, custom controls, and destinations other than built-in Frames or
Pages. Changing a direct Page parent relationship also requires the Page edit
capability. Canonical validation rejects invalid references and
cycles atomically. Generic Go edits retain their broader existing contract.

A clickable breadcrumb shows the form and selected control's ancestors.
Clicking the form or an ancestor selects it without editing source. New UI
strings use the extension localization channel. Selection, destination
highlighting and preview geometry are transient state. A stale document or
connection, Escape, pointer cancellation or capture loss discards the gesture.

See [ADR-0069](../adr/ADR-0069-userform-designer-frame-parenting.md).

## Property Grid and authored source states

The running LSP supplies `propertyGrid` alongside the canonical preview
document. The form and each explicit control ID have a property target:

```json
{
  "propertyGrid": {
    "form": {
      "descriptors": [
        { "field": "caption", "valueType": "string", "required": false, "nullable": true }
      ],
      "values": { "caption": { "present": false, "value": null } }
    },
    "controls": {
      "submit": {
        "descriptors": [
          { "field": "enabled", "valueType": "boolean", "required": false, "nullable": true }
        ],
        "values": { "enabled": { "present": true, "value": false } }
      }
    }
  }
}
```

This example shows a subset of descriptors. Each descriptor contains `field`,
`valueType`, `required`, `nullable` and optional `allowedValues`. Field paths
are relative to the form or selected control. Descriptors derive from the Go
FormSpec contract and scalar edit support for that target; TypeScript owns
labels, categories and display order, not another validation/applicability
schema. The UI supports string, number, integer, boolean and enum inputs;
scalar `value` retains its chosen type, including string `"false"` versus
boolean `false`.

`values` comes from original source syntax, not the normalized document or
renderer. `{present: false, value: null}` means absent;
`{present: true, value: null}` means explicit null. Empty string, false and zero
remain explicit values. Observed values and approximate display defaults are
never used to fill or persist authored fields. A preview-valid source whose
authored state cannot be determined safely returns optional `propertyError`
(using the preview error shape) without discarding its valid `document`.
Property editing is unavailable without usable metadata.

Authored numeric values must survive the JavaScript number wire boundary without
changing their decimal value. Metadata that would round a large integer,
high-precision decimal, or underflowing value is unavailable with a structured
property error; the render preview remains available. Numeric input applies the
same decimal round-trip check before sending an edit and retains rejected drafts.
Ordinary decimal and exponent spellings (for example `0.1` and `1e2`) remain
editable. These restrictions apply to the Property Grid, not canonical FormSpec
source validation or source-editor edits.

Background selection shows form fields; control selection shows that control's
fields. Selecting a Page by its stable ID in the Page list or by a rendered
header shows that Page's metadata fields and excludes Page geometry. Page
preview navigation is transient; explicit MultiPage and TabStrip
`selectedIndex` edits continue through this Property Grid. Form fields include
`name`, `caption`, `width`, `height` and the supported `build.*` paths;
legacy and build fields are separate and never redirect to each other.
Control fields are limited to supported authored scalars for their type.
Valid custom ProgID controls retain canonical common scalar fields, without
acquiring built-in type-specific fields.
Observed/snapshot values, identity/type/parent topology, collections, pictures
and arbitrary property bags are outside the scalar grid. Dedicated Page and
TabStrip editors manage those collections using their separate operations.
Changing `form.name` does not rename the source file, code sidecar or VBA
references.

Text/number inputs confirm on Enter or blur, with at most one request when both
occur; Escape cancels the draft. Boolean/enum changes confirm on selection.
Keep invalid input with a field-level error without changing source; preserve
structured field/code/message diagnostics from the server. Drafts belong to
target kind, target ID, field and document version. Form identity is distinct
from every control ID, including a control literally named `form`, in both
component keys and draft/reply storage. External source edits invalidate
unsubmitted Property Grid field drafts rather than rebasing those field values.
The URI-shared source draft follows the deferred-save merge and conflict policy
above. Switching selection preserves dirty Property Grid field drafts and
their matching edit replies, including field-level rejection diagnostics,
until corrected or canceled. Source edits and
undo/redo refresh values from source. Selection and rendering alone never
produce source edits. Explicit null is allowed only for nullable fields
accepted by canonical validation; there is no property removal action.

The host retains edit results while the panel is hidden and replays them when
the same Webview becomes visible. A failed message delivery keeps its result
queued for the next replay. A new Webview `ready` starts a new request context:
queued results from the old context are discarded, and old in-flight edits
cannot apply or acknowledge new requests whose IDs restart from one. Visibility
still participates in the queue's stale-edit guard; retaining an error reply
never authorizes a hidden or otherwise obsolete source edit.

## Page and TabStrip editors

Selecting a MultiPage exposes its Page collection in the property region. The
list identifies Pages by stable control ID and includes Pages with hidden
headers. Clicking a Page changes transient `activePageId` and the preview; it
does not issue an edit or persist `selectedIndex`. Visible Page headers provide
the same navigation. Page `name` and `caption` remain editable through the
existing LSP property metadata. Page geometry is derived and is never exposed
as an authoring field.

The Page list provides Add, Delete, and explicit move-up/move-down actions.
Adding appends a Page to the selected MultiPage. Deleting a Page with descendants requires
confirmation that names the Page and descendant-control count, then removes
the Page and its descendants with `cascade: true` in one edit. Empty Pages are
deleted immediately. Reordering
changes order only within that MultiPage; hover-tab drag does not reorder.
Controls can be inserted into the active Page, dropped on a Page, or assigned
to a Page by the destination picker. The picker changes ordinary controls'
parent. A Page control cannot be reparented to another MultiPage.

Selecting a TabStrip exposes a separate tab editor. It lists every record,
including tabs with `visible: false`, and edits `name`, `caption`, `enabled`,
`visible`, `tag`, `controlTipText`, and `accelerator` using field metadata and
authored values from the LSP. Tabs remain records in `tabs`; they are not
controls or Pages. TabStrip `selectedIndex` remains a control property edited
through the Property Grid.

Tab property drafts use the owning control ID and canonical unique tab name,
separately from the current index used by the edit protocol. Reordering therefore
keeps each tab's own draft; submitting resolves the current index. Source version
changes still mark unsubmitted drafts stale until the user edits or cancels them.

The Page operations are `addPage(control)`, `removePage(controlId, cascade)`,
and `reorderPage(controlId, index)`. Tab operations are
`addTab(controlId, tab)` (append), `removeTab(controlId, index)`,
`reorderTab(controlId, index, toIndex)`, and
`setTabProperty(controlId, index, field, value)`. Each user action creates one
semantic edit and one native Undo/Redo step.

`userFormPageEdit` independently gates Page collection operations and Page
destinations. `userFormTabEdit` independently gates TabStrip collection and
tab-property operations. Editing Page `name`/`caption` or control-level
`selectedIndex` uses `userFormPropertyEdit`. Inserting into a Page also needs
`userFormStructuralEdit` and `userFormParentEdit`; changing a control's Page
parent by drag or the destination picker also needs the existing geometry
capability for `setParent` coordinates. Missing capabilities disable only their
corresponding actions.

Collection edits preserve selected-member identity when an authored numeric
`selectedIndex` exists. Reordering changes the index to retain that member;
deleting it selects index zero if entries remain, or logical `-1` for an empty
collection. Adding the first member changes an explicit `-1` to zero. Absent
and explicit-null values remain unchanged. Page navigation alone never writes
`selectedIndex`, and observed or normalized values never become authored
values. YAML/YML/JSON edits are localized and preserve unrelated source. The
server validates the final FormSpec after applying the transaction; a failure
returns no partial edit. These operations do not change Excel, COM, VBIDE, or
runtime semantics.

## Parent-local tab order

Provide an editor for each parent whose direct controls expose editable
`tabIndex` metadata from the LSP. Include hidden and disabled controls. Exclude
Page controls and TabStrip tab records. The root form uses `parentId: ""`;
Frame and Page lists use their stable control IDs. Up/down changes are local
draft state; Apply sends the complete ordered list of eligible sibling IDs and
assigns contiguous zero-based `tabIndex` values in one source edit.
When LSP metadata yields no eligible controls for the selected parent, show the
localized `noEditableTabOrder` empty-state string.

Initialize the list from explicitly authored numeric `tabIndex` values in
ascending order. Ties follow stable z-order, preserving source order for
z-order ties. Missing and explicit-null indexes follow numeric entries in
z-order. Observed or normalized indexes do not become authored values. If the
source document changes before Apply, discard the local order and rebuild it
from the current document; do not merge or rebase it.

## Semantic edit protocol and native history

The LSP advertises `capabilities.experimental.userFormEdit: true`. Preview-only
servers retain read-only rendering, selection, grid and zoom. Requests use
`xlflow/userFormEdit` with `{uri, version, text, operations}`. Operations are
`moveControl` (both `left`/`top`) and `resizeControl` (both `width`/`height`),
using control IDs and finite point values, or one `setFormProperty` /
`setControlProperty` operation. Property editing additionally requires
`capabilities.experimental.userFormPropertyEdit: true`; older geometry-capable
servers retain geometry editing with the grid disabled. Preview-only servers
retain read-only rendering. The extension and RPC both accept
one operation, a move/resize pair for the same control, or a distinct numeric
`setFormProperty` width/height pair per interaction. A completed form corner
resize updates both dimensions in one source transaction and one Undo step.
Adding and removing ordinary controls additionally require the separate
`capabilities.experimental.userFormStructuralEdit: true` capability. This
capability gates `addControl` and `removeControl` without changing the
independent geometry and property capability checks. Page collection operations
require `capabilities.experimental.userFormPageEdit: true`; TabStrip collection
and tab-property operations require
`capabilities.experimental.userFormTabEdit: true`. These capabilities are
independent of one another and of `userFormPropertyEdit` and
`userFormParentEdit`. An older server that omits a capability keeps its other
advertised edit families and does not expose the unsupported actions.
Each Toolbox insertion or deletion sends one structural operation; it is not
combined with geometry or property operations in the same transaction.
The RPC also enforces structural scope: `addControl` accepts a root, Frame, or
eligible Page destination, but not a Page control. A Page destination requires
structural, parent, and Page capabilities. An omitted or empty ProgID uses the
built-in type contract; an explicit nonempty ProgID must match that contract's
ProgID case-insensitively, without trimming whitespace. Custom ProgIDs are
rejected before source edits are generated. `removeControl` continues to reject
a Page target; `removePage` is the dedicated Page deletion operation and
requires Page capability. Page targets resolve through the canonical flattened
source model, including legacy nested input. Deleting a MultiPage with
`cascade: true` may still remove its Pages and their descendants. The shared
semantic edit API retains generic valid Page relationships and collection
operations; the Designer adapter enforces its capability and interaction
scope.
For existing operation families, empty batches, more than two operations,
repeated operation types, and pairs targeting different controls are rejected
before source edits are generated, except for the dedicated width/height form
property pair. Other property pairs and mixed property/geometry edits remain
rejected; duplicate dimensions, nulls and nonnumeric values in a form resize
pair are invalid.
Its payload supplies `field` and an explicit scalar `value`, plus `controlId`
for `setControlProperty`. Missing `value` is invalid; explicit null must survive
transport as a present value. Both Webview boundary and RPC validate payload
shape, while canonical parsing/validation decides FormSpec semantics.

The server validates eligibility and calls `spec/edit.Apply` on the supplied
unsaved text. It returns `{version, edits, warnings}` or `{version, error}`,
with standard LSP TextEdit ranges converted from UTF-8 offsets to UTF-16
positions. Errors retain structured operation diagnostics and canonical codes.
The server does not apply edits, write files or replace input with saved source.

The host serializes edits per document URI, captures draft source/revision,
validates returned ranges, and applies each completed action to the shared
draft as one `CustomDocumentEditEvent`. Reject requests and responses obsolete
after draft, project, configuration, connection, visibility, or disposal
changes. Native custom-document history owns undo/redo; there is no Webview
history. External text-editor changes refresh the draft or enter the explicit
conflict state described above. Invalid source or unavailable
connections disable editing while retaining the last valid rendering. Rejected
operations revert preview and display an error without automatic retry/rebase.

Webview messages add `edit` with request ID/version/operations. Host messages
add edit results and availability; document messages carry editable capability.
YAML, YML, and JSON support the semantic operations in
[`userform-semantic-edit.md`](userform-semantic-edit.md), including JSON
`reorderControl`, while preserving unrelated source formatting. Canvas form
resizing remains outside this contract.
Frame and Page placement/reparenting use the separate parent and Page
capabilities described above. Page and TabStrip operation details and
`selectedIndex` identity rules are specified in
[`userform-semantic-edit.md`](userform-semantic-edit.md) and
[`userform-multipage-tabstrip.md`](userform-multipage-tabstrip.md).
Structural control add/remove continues to use the separate capability above.

The LSP advertises `capabilities.experimental.userFormProductivityEdit: true`
independently of existing capabilities. Each `userFormEdit` request containing
`moveControls`, `alignControls`, `distributeControls`, `sameSizeControls`,
`reorderControls`, `removeControls`, `setTabOrder`, or `pasteControls` contains
exactly that one operation. Do not combine a productivity operation with a
legacy operation or another productivity operation. Each successful
productivity action is applied by the extension host as one
`CustomDocumentEditEvent`.
The new capability does not imply another edit family: movement, layout, and
z-order also require `userFormEdit`; `setTabOrder` also requires
`userFormPropertyEdit`; `removeControls` and `pasteControls` also require
`userFormStructuralEdit`. A paste that sets parent relationships additionally
requires `userFormParentEdit`; creating or placing Page controls requires
`userFormPageEdit`. The host and LSP enforce all applicable gates separately.
Older servers keep only the other edit families they already advertise.

The extension host sends `xlflow/userFormCopy` with
`{uri, version, text, controlIds}` using the current unsaved custom-document draft. The
request requires `userFormProductivityEdit`. A successful response is
`{version, snapshot, document, hasPicturePaths:boolean}`. An error response
contains `{version, error}` and may also contain `hasPicturePaths:boolean`; that
field has no meaning without a successful snapshot. `hasPicturePaths` is
required on success, including when false. `ClipboardSnapshot` is
`{format: 'yaml'|'json', text, controlIds}`. The `controlIds` and copied root
records follow stable sibling stacking order regardless of selection order; authored
`zIndex` fields remain source data. The backend determines the exact
same-source records in `text`, preserving source presence and observations and distinguishing
absent, null, empty, false, and zero without synthesizing normalized defaults.
Stale source/version results are discarded.

A successful Copy publishes its snapshot to the shared clipboard. Cut keeps
its captured snapshot provisional through confirmation and deletion, and
publishes it only after the atomic removal succeeds.

Each Copy or Cut selection request receives a global monotonically increasing
ticket. Only the latest valid requested selection may publish its snapshot;
discard an older response that arrives late. A Cut must still hold the current
ticket before cascade deletion; if a newer request supersedes it first, reject
the Cut and leave its source unchanged. Replacing the clipboard snapshot resets
paste repetition offsets for every destination URI/parent pair. A paste may
finish with the snapshot it captured before a replacement, but it must not
advance the replacement snapshot's offsets.

The host proposes one new stable ID and case-insensitively unique name for every
copied control record, including descendants, and sends
`pasteControls{snapshot,allocations:[{sourceId,id,name}],parentId,offset:{left,top}}`.
The LSP rewrites internal IDs and parent references, assigns copied roots to
the requested parent, and validates the final FormSpec atomically. The host
adds an 8 pt offset on both axes for each repeat and applies only a uniform
translation to fit the complete group; it rejects a group that cannot fit.
Page is excluded as a direct copy root, while a copied MultiPage includes its
Page subtree. Unsupported custom subtrees fail as a whole. Same-project
project-root-relative picture paths remain unchanged; cross-project paste
rejects a subtree containing one. Image files, event handlers, and code-behind
are not copied. Cut captures a valid snapshot, confirms the aggregate
descendant count in the same URI queue, then submits one atomic `removeControls`
edit with cascade enabled. Only successful deletion publishes the captured
snapshot; cancellation or failure preserves the source and previous clipboard.
Duplicate uses the same snapshot path and one paste operation in the original
selection's owning parent. Paste uses a selected Frame or Page as destination
only for a single selected control; otherwise it uses the selection's owning
parent.

The tab-order editor is parent-local and lists only controls whose LSP metadata
exposes editable `tabIndex`. Hidden and disabled controls remain listed; Page
controls and TabStrip tab records are excluded. It initializes explicit
numeric indexes ascending, ties by z-order, and absent/null indexes last by
z-order. Apply sends `setTabOrder{parentId,controlIds}` with the complete
ordered list; the LSP assigns zero-based indexes in one edit. A source change
discards a pending local order instead of rebasing it.

## Deferred-save verification requirements

These are required implementation checks; they do not claim that the current
implementation or test run has passed.

Verify that multiple Designer panels for one URI share draft text and native
Undo/Redo. A completed LSP edit must change only the draft until explicit
Designer Save or its scoped keyboard command. Native Save, Save All, and
configured AutoSave must checkpoint internally without writing the open
`TextDocument`; Undo/Redo must not publish it. Verify that the tab dirty marker
tracks the internal checkpoint independently from the YAML/JSON-not-applied
badge.
Preview and edit requests must carry the exact draft text, preserve source
formatting, and retain canonical LSP validation behavior without frontend
semantic rules.

Verify that external unsaved text-editor changes appear immediately. A clean
draft adopts them; non-overlapping changes merge into a dirty draft and the
EngineWorker rebases history so subsequent Undo/Redo remains correct. An
overlapping edit must retain both versions, show the Source Conflict state, and
block explicit Designer Save until an explicit resolution. Native checkpoint
save and AutoSave remain available without publishing source. Cover **Use YAML/JSON
changes**, **Use Designer draft**, and an explicitly confirmed whole-document
choice when a granular merge is unavailable.

Verify explicit Designer Save validates and applies edits to the latest open
`TextDocument`, then calls its save API without writing the workspace file
directly or replacing newer dirty text. Verify the native save callback
checkpoints only internal storage for menu Save, Save All, and AutoSave, and
that pending drafts survive close/reopen/restart by URI. Cover failed
apply/save retry, Save As as an explicit publish, Revert, backup, restore,
hot-exit, and standard AutoSave enabled/disabled settings. A corrupt checkpoint
must show an error and remain stored until successful explicit Save or Revert
resolves the pending recovery state.

## Structural edit verification requirements

Implementation checks should cover the 14 Toolbox types and Pointer, click-to-
place and Escape cancellation, centralized default dimensions, stableDelta
precision and grid snapping, position clamping and undersized-form rejection,
UUID/name collision handling at Toolbox allocation and host add preflight,
stable ID after rename, and the empty MultiPage/TabStrip states. Designer checks
should cover root, Frame, and active-Page insertion, selection followed by
return to Pointer, canvas Delete handling that leaves Pages intact, Page-list
deletion confirmation and cascade, immediate leaf-control deletion,
confirmation content and cancellation for controls with descendants, descendant
cascade, a single native Undo/Redo step, and pending-confirmation/version
invalidation.

Keep coverage in `editors/vscode/test/userFormStructure.test.tsx`, the actual
Webview interaction suite `editors/vscode/test/suite/userFormWebview.test.ts`,
and host preparation checks in `editors/vscode/test/suite/userFormProvider.test.ts`.

Go and LSP checks should cover JSON add/remove/parent source edits, source formatting
preservation, removal of descendants represented by legacy nested source, and
final canonical validation without a new global name-uniqueness rule. Verify
that older servers retain their separately advertised geometry/property
behavior and do not expose structural actions without
`userFormStructuralEdit`. These checks concern source editing and do not claim
Excel, COM, VBIDE or VBE behavior.

## Property editing verification

Go/LSP checks cover per-type descriptors, absent/null/empty/false/zero states,
Page geometry exclusion, ambiguous-source `propertyError` with valid preview,
canonical rejection, value presence, transaction shape and capability fallback.
Source tests cover YAML/YML/JSON local insertion/replacement, missing `build`,
CRLF and multibyte offsets, escaped/duplicate JSON keys, no-ops and atomic
failure. Webview checks cover selection, typed input, Enter/blur deduplication,
Escape, invalid draft retention and draft invalidation on source updates.
Verify in real VS Code that one confirmation creates one native undo/redo
step, multiple panels synchronize, stale versions/connections cannot overwrite
source and older servers retain their advertised behavior. These source edits
do not change COM/VBE semantics and require no Excel or VBE oracle run.

## Page and TabStrip verification requirements

Designer checks should verify that Page-list and rendered-header selection use
Page IDs, including hidden Pages, and do not edit source or `selectedIndex`.
Cover Page add, confirmed cascade deletion and cancellation, up/down reorder,
active-Page insertion/drop, the Page destination picker, rejection of Page
reparenting, and exclusion of Page geometry. TabStrip checks should list hidden
and visible records and cover append, remove, reorder, and LSP-metadata-driven
editing of every supported field. Source/LSP checks should cover explicit
selection identity through reorder/delete, fallback zero and empty `-1`, first
addition after explicit `-1`, retained absent/null values, independent
capability gates, localized YAML/YML/JSON edits, observed-data preservation,
and atomic final-validation failure. These checks do not establish runtime,
Excel, COM, or VBE behavior.

## Productivity verification requirements

Plan implementation coverage in three layers. First, run the focused Go
semantic-edit and LSP checks through `scripts/dev/go.ps1` for YAML, YML, and
JSON, including JSON `reorderControl`, one-operation transaction enforcement,
source preservation, stale copy versions, and atomic rejection. Cover
selection scoping, movement, six alignments, fixed-endpoint distribution,
anchor sizing, all four z-order commands, complete tab-order lists, and
recursive copy/paste rewriting and allocation.

Second, run the VS Code Designer tests for built-in-only group selection,
standalone custom-control selection, marquee inclusion of fully contained
visible built-in siblings, group movement without resizing, context actions,
Cut snapshot-confirm-delete ordering and cancellation preservation,
single-container Paste and owning-parent Duplicate destinations, cross-panel
clipboard use, same-project and cross-project picture paths, and tab-order
initialization, Apply, and discard after source changes. Include malformed,
unsupported custom, direct-Page, and out-of-bounds copy/paste cases.

Third, verify in the VS Code host that each successful productivity action is
one native Undo step and Redo restores the complete action. Exercise cut,
paste, alignment, distribution, z-order, and tab-order separately. These are
source/editor checks; do not report them as Excel, COM, VBIDE, VBE, compiler,
or runtime validation.

## Webview security

Each Webview document uses a fresh random nonce for its bundled script. Its
Content Security Policy sets `default-src 'none'`, permits scripts only with
that nonce, and permits styles only from the VS Code Webview resource origin.
The Webview's local resource root is limited to the packaged Designer bundle.
Workspace-provided strings are passed as data to the renderer and are not
interpolated into the HTML document.

## Related

- Issue #915
- Issue #917
- Issue #918
- Issue #919
- Issue #920
- Issue #921
- Issue #922
- `docs/adr/ADR-0068-userform-designer-structural-control-edits.md`
- `docs/adr/ADR-0067-userform-designer-property-metadata.md`
- `docs/adr/ADR-0066-userform-designer-geometry-transactions.md`
- `docs/adr/ADR-0073-userform-designer-deferred-save-document-lifecycle.md`
- `docs/adr/ADR-0070-userform-designer-page-tab-collection-editing.md`
- `docs/adr/ADR-0071-userform-designer-productivity-edits.md`
- `docs/adr/ADR-0064-userform-designer-text-document-boundary.md`
- `docs/specs/ms-oforms.md`
- `docs/specs/userform-picture-assets.md`
