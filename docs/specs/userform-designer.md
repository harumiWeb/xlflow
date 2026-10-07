# VS Code UserForm Designer

This specification defines the interactive UserForm Designer in the
xlflow Visual Studio Code extension. The VS Code `TextDocument` remains the
canonical source. The Designer renders the current document through the xlflow
Language Server Protocol (LSP); it does not create or persist another source
representation.

## Editor selection and document eligibility

The extension registers `xlflow.userFormDesigner` as an optional custom text
editor for `.yaml`, `.yml`, and `.json` files. Its `option` priority keeps the
normal text editor available. Users can choose the visual editor with **Reopen
Editor With...**.

The preview is available only for a FormSpec directly inside the `specs`
directory under the configured forms root (`[src].forms` in `xlflow.toml`,
defaulting to `src/forms`). Arbitrary YAML, YAML nested in another directory,
and other JSON files are not UserForm documents. The file must belong to an
xlflow project with `xlflow.toml` and be handled by the currently selected
project's single xlflow LSP client. Opening a document associated with another
project, or without a compatible LSP connection, produces an editor error.

## Source and preview protocol

The extension host reads source from the open VS Code `TextDocument`, including
unsaved changes. It sends the document URI, its current version, and the full
text to the LSP. The Webview never reads the document or workspace files
directly.

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
as an independently positioned top-level control. The selected Page from the
FormSpec is displayed, defaulting to the first Page when no selection is
recorded.

Explicitly invisible controls and their descendants are not rendered. Invisible
Page tabs are omitted, but selection indexes retain the full ordered Page
collection. A selected invisible Page does not fall through to a different
Page. Omitted visibility defaults to visible.

Page collection order follows stable sibling `zIndex` order, matching the
pure-Go generation and topology compiler. Selection is applied before hiding
invisible Pages. Explicit build client dimensions take precedence over observed
client/inside dimensions, which take precedence over outer dimensions. Either
axis relying on outer dimensions remains approximate. ComboBox display prefers
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

Selection is a transient control ID; clicking the background selects the form.
Known built-in controls other than Page can move and resize with eight handles.
Form and Page geometry is not edited on the canvas. Unknown controls remain
selectable placeholders. No reparenting or structural edit is implicit.

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
is never fixed on load; fitting moves/resizes are permitted, but an oversized
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

## Toolbox and structural control edits

The Toolbox contains Pointer and these 14 built-in control types: Label,
TextBox, ComboBox, ListBox, CommandButton, CheckBox, OptionButton,
ToggleButton, SpinButton, ScrollBar, Image, Frame, MultiPage, and TabStrip.
Page is not a Toolbox item. Pointer is the normal selection tool. Selecting a
control type arms placement; clicking the canvas inserts one control at the
click position. Drag-to-create is not part of this interaction.

New controls are always root controls. The placement UI does not target an
existing container or assign `parentId`, and it does not infer containment
from overlap. Use the same centralized per-type width and height defaults listed under Rendering;
creation and preview must not maintain separate size tables. Convert the
canvas click to FormSpec points. With snapping off, normalize each coordinate
with the TypeScript `stableDelta` helper (six decimal places) to suppress
floating-point micro-error; with snapping on, use the existing 8 pt grid. Then
clamp `left` and `top` so the default-size control fits within the current form
bounds. If either form dimension is smaller than the new control's default
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
A control with no descendants is removed immediately. If descendants exist,
the host first shows a modal confirmation containing the selected control's
name and descendant count. Confirmed removal deletes the selected control and
all descendants with `cascade: true` in one semantic edit and one VS Code
`WorkspaceEdit`; Undo/Redo therefore restores or removes the whole subtree as
one native history step. Direct deletion of a selected Page is disabled.

Structural edits apply only to the open FormSpec `TextDocument`. They do not
edit code-behind or save the document automatically. The host captures the
source and version associated with an add request or pending delete
confirmation. If the document text/version, project, or LSP connection changes
before the edit is submitted or applied, discard the stale action and refresh
from the current document; do not replay it against newer text.

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
fields. Page selection uses its tab and excludes Page geometry. Form fields
include `name`, `caption`, `width`, `height` and the supported `build.*` paths;
legacy and build fields are separate and never redirect to each other.
Control fields are limited to supported authored scalars for their type.
Valid custom ProgID controls retain canonical common scalar fields, without
acquiring built-in type-specific fields.
Observed/snapshot values, identity/type/parent topology, collections, pictures
and arbitrary property bags are outside the grid. Changing `form.name` does
not rename the source file, code sidecar or VBA references.

Text/number inputs confirm on Enter or blur, with at most one request when both
occur; Escape cancels the draft. Boolean/enum changes confirm on selection.
Keep invalid input with a field-level error without changing source; preserve
structured field/code/message diagnostics from the server. Drafts belong to
target kind, target ID, field and document version. Form identity is distinct
from every control ID, including a control literally named `form`, in both
component keys and draft/reply storage. External source edits invalidate old
drafts rather than applying or rebasing them automatically. Switching selection
preserves dirty drafts and their matching edit replies, including field-level
rejection diagnostics, until corrected or canceled. Source edits and
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
one operation or a move/resize pair for the same control per interaction.
Adding and removing controls additionally require the separate
`capabilities.experimental.userFormStructuralEdit: true` capability. This
capability gates `addControl` and `removeControl` without changing the
independent geometry and property capability checks. An older server that does
not advertise it keeps the preview and whichever existing edit capabilities it
advertises; structural actions are unavailable.
Each Toolbox insertion or deletion sends one structural operation; it is not
combined with geometry or property operations in the same transaction.
The RPC also enforces this structural scope: insertion accepts only a root
built-in control other than Page, and direct Page deletion returns
`unsupported` without edits. Page targets resolve through the canonical
flattened source model, including legacy nested input. Deleting a MultiPage
with `cascade: true` may still remove its Pages and their descendants. These
restrictions belong to the Designer adapter; the shared semantic edit API
retains its general parent and Page operations.
Empty batches, more than two operations, repeated operation types, and pairs
targeting different controls are rejected before source edits are generated.
A property operation cannot be batched with another property or geometry edit.
Its payload supplies `field` and an explicit scalar `value`, plus `controlId`
for `setControlProperty`. Missing `value` is invalid; explicit null must survive
transport as a present value. Both Webview boundary and RPC validate payload
shape, while canonical parsing/validation decides FormSpec semantics.

The server validates eligibility and calls `spec/edit.Apply` on the supplied
unsaved text. It returns `{version, edits, warnings}` or `{version, error}`,
with standard LSP TextEdit ranges converted from UTF-8 offsets to UTF-16
positions. Errors retain structured operation diagnostics and canonical codes.
The server does not apply edits, write files or replace input with saved source.

The host serializes edits per document URI, captures source/version, validates
returned ranges and applies all edits in one WorkspaceEdit. Reject requests and
responses obsolete after text, project, configuration, connection, visibility
or disposal changes. Native custom text editor history owns undo/redo; there
is no Webview history or automatic save. Text edits, undo/redo and other panels
refresh through the preview synchronizer. Invalid source or unavailable
connections disable editing while retaining the last valid rendering. Rejected
operations revert preview and display an error without automatic retry/rebase.

Webview messages add `edit` with request ID/version/operations. Host messages
add edit results and availability; document messages carry editable capability.
YAML/YML and JSON share geometry and scalar property interactions, preserving
unrelated source formatting. Multiple selection and canvas form resizing remain
outside this contract. Root-only insertion does not include placement inside a
Frame or reparenting (#920); Page and Tab collection editing remain out of
scope (#921).
Structural add/remove use the separate capability above.

## Structural edit verification requirements

Implementation checks should cover the 14 Toolbox types and Pointer, click-to-
place and Escape cancellation, centralized default dimensions, stableDelta
precision and grid snapping, position clamping and undersized-form rejection,
UUID/name collision handling at Toolbox allocation and host add preflight,
stable ID after rename, and the empty MultiPage/TabStrip states. Designer checks
should cover root-only insertion, selection followed by return to Pointer,
disabled direct Page deletion, canvas-only Delete handling, immediate leaf
deletion, confirmation content and cancellation for controls with descendants,
descendant cascade, a single native Undo/Redo step, and pending-confirmation/
version invalidation.

Keep coverage in `editors/vscode/test/userFormStructure.test.tsx`, the actual
Webview interaction suite `editors/vscode/test/suite/userFormWebview.test.ts`,
and host preparation checks in `editors/vscode/test/suite/userFormProvider.test.ts`.

Go and LSP checks should cover JSON add/remove source edits, source formatting
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
- `docs/adr/ADR-0068-userform-designer-structural-control-edits.md`
- `docs/adr/ADR-0067-userform-designer-property-metadata.md`
- `docs/adr/ADR-0066-userform-designer-geometry-transactions.md`
- `docs/adr/ADR-0064-userform-designer-text-document-boundary.md`
- `docs/specs/ms-oforms.md`
- `docs/specs/userform-picture-assets.md`
