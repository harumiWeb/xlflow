# Manage UserForms

Use `form snapshot` to capture Designer state, `form build` to create from a persisted specification, and `form export-image` to verify the rendered result. In sidecar mode, Designer structure lives in `src/forms/specs/` and code-behind in `src/forms/code/`.

```bash
xlflow form snapshot UserForm1 --json
xlflow form build src/forms/specs/UserForm1.yaml --json
xlflow form export-image UserForm1 --json
```

Run `lint` and `form build` preflight before opening Excel. Do not edit generated `.frm` code-behind independently when sidecar mode is authoritative; see the [UserForm specification](../reference/userform-spec).

Root `form.width` and `form.height` author persisted MS-OFORMS `DisplayedSize`
in points. Snapshots and file pull read that same persisted value; VBFrame
`ClientWidth` / `ClientHeight` are compiler-maintained companion values and do
not seed authoring dimensions. Excel import/export may normalize `DisplayedSize`,
so a snapshot reports the persisted result and does not guarantee exact equality
with the authored points. This Excel behavior is separate from pure-Go
HIMETRIC rounding; xlflow does not compensate with a fixed offset.

## Visual layout in VS Code

Open a canonical `.yaml`, `.yml`, or `.json` FormSpec directly under your
configured forms root's `specs` directory, then use **Reopen Editor With... →
xlflow UserForm Designer**. Select a control to drag it or resize with one of
eight handles. Arrow keys move 1 point; Shift+Arrow moves 10 points. Holding
an arrow key creates one edit when the keys are released. Escape cancels an
unfinished interaction. Use VS Code Undo/Redo to restore completed edits.

Select the form by clicking its background or its name in the selection path.
Drag the right edge, bottom edge or bottom-right corner to resize it. This
updates `form.width` and `form.height` in client-area points; the title bar and
window border are not included. A corner drag updates both dimensions in one
Undo step. Shrinking stops at the extents of existing root controls.

The grid displays 8-point spacing and has a separate snapping toggle. Snapping
applies to mouse movement and resize; keyboard movement retains its 1/10-point
steps. Zoom ranges from 50% to 200% and does not change source coordinates.
These settings and selection are temporary Designer state.

Movement and resize stay inside the parent's displayed content area. Shrinking
a Frame or MultiPage stops before its children would overflow, including
children on other Pages. Existing out-of-bounds controls are not automatically
corrected; controls larger than their parent must be resized to fit before
moving. Direct source edits remain available for layouts outside these UI
constraints. Page bounds are not directly editable on the canvas.

Visual operations produce local source edits and retain unrelated formatting.
Invalid source retains the last valid preview and disables editing until fixed.
The Toolbox contains Pointer and Label, TextBox, ComboBox, ListBox,
CommandButton, CheckBox, OptionButton, ToggleButton, SpinButton, ScrollBar,
Image, Frame, MultiPage, and TabStrip. Choose a type, then click the canvas to
insert a root control or target a Frame; Escape cancels placement. The Designer uses the same
type-specific default sizes as its preview and clamps the position to fit the
destination. A destination smaller than the default control size rejects insertion. New
controls have stable `control-<UUID>` IDs and conflict-free generated VBA names.
Insertion selects the new control and returns to Pointer. A new MultiPage has
no Pages and `selectedIndex: -1`; a new TabStrip has `tabs: []` and
`selectedIndex: -1`.

Delete the selected control with the toolbar or Delete while the canvas has
focus. Leaf controls are removed immediately. When a control has descendants,
a confirmation shows its name and descendant count; confirming removes the
subtree as one undoable edit. Designer deletion does not intercept Delete in a
focused input, so normal character deletion continues to work. Direct deletion
of a Page is disabled. Structural editing requires the LSP's separate
`userFormStructuralEdit` capability. Older servers keep their advertised
geometry/property support and do not offer structural actions. Source edits do
not change code-behind or save automatically.

Drag an existing control into a Frame or onto the form background to change
its parent. The highlighted destination and drag preview show the final
position before dropping. Coordinates are converted to the new parent's
space, snapped when enabled, then clamped to its content area. A control too
large for the destination cannot be dropped. Moving a Frame carries its
children; resizing and keyboard movement retain the current parent. Use the
clickable selection path above the canvas to select ancestor Frames.

Frame placement additionally requires `userFormParentEdit`; reparenting
requires both that capability and geometry support. Older servers retain
root placement and same-parent movement. Page and TabStrip collection editing
is available when the LSP advertises `userFormPageEdit` and
`userFormTabEdit`, respectively. See the sections below for Page placement,
tab editing, and tab order.

### Selection and layout productivity

Ctrl/Cmd-click or Shift-click toggles ordinary built-in sibling controls in a
group selection under the form, one Frame, or one Page. Custom controls remain
selectable alone but cannot join a group. Marquee selection includes only
visible, ordinary built-in immediate children whose complete displayed bounds
are inside the marquee. Dragging a group preserves its parent and relative
geometry; group dragging does not resize controls.

Use Align Left/Center/Right or Top/Middle/Bottom to align to the last selected
control. Make Same Width, Same Height, or Same Size also uses that anchor.
Distribute Horizontally or Vertically requires at least three controls and
keeps the first and last positions fixed while making the edge gaps equal. The
Designer sorts by left/top position and uses z-order to break ties.
Bring to Front, Bring Forward, Send Backward, and Send to Back preserve the
selected controls' relative order and leave `tabIndex` unchanged.

The selection context menu offers Cut, Copy, Paste, Delete, Duplicate, z-order
commands, and Properties. Copy publishes its successful snapshot to an
in-memory clipboard shared by Designer panels. Copying a Frame includes its
children; copying a MultiPage includes its Pages and their descendants. A Page
cannot be copied directly, and unsupported custom-control subtrees are
rejected. Paste uses a selected Frame or Page as its destination only when it
is the sole selected control; otherwise, including a multi-selection, it uses
the selection's owning parent. Duplicate always uses the original selection's
owning parent, so a Frame is not duplicated inside itself. Each repeated paste
starts 8 points down and right; a new clipboard snapshot restarts this offset
for each destination form and parent. The Designer shifts the entire pasted
group to fit its destination or rejects it if it cannot fit without resizing.
Same-project picture paths stay unchanged; a cross-project paste containing
one is rejected. Image files, events, and code-behind are not copied.

Cut first captures a valid source snapshot, then asks one confirmation for the
aggregate descendant count across selected controls; leaf-only Cut needs no
descendant confirmation. After confirmation, it removes the selection as one
undoable edit and publishes the captured snapshot to the shared clipboard only
if deletion succeeds. Cancellation or failure leaves the source and previous
clipboard unchanged. Each completed
layout, z-order, paste/duplicate, delete/cut, or tab-order Apply action is one
VS Code Undo step. Copy alone does not edit the document. Productivity actions
require the LSP's `userFormProductivityEdit` plus the applicable existing
geometry, property, structural, parent, or Page capability.

### Edit properties

Select the background for form properties, a control for its properties, or a
Page tab for Page properties. The Property Grid shows supported scalar fields
from the running LSP. Form legacy fields and `build.*` fields are separate;
changing `form.name` does not rename the file, code sidecar or VBA references.
Page geometry, observed fields, lists/tabs, pictures and property bags are not
edited in this grid.

Confirm text/number input with Enter or by leaving the field; Escape cancels
the draft. Boolean/enum choices confirm immediately. Each confirmation creates
one native undoable source edit in YAML, YML or JSON. Invalid input keeps its
field error without changing source. A source update invalidates an older
draft; re-enter it against the refreshed document.

Unset, explicit null, empty string, false and zero are distinct. The grid does
not fill authored fields from measured or preview defaults. Explicit null is
available only for supported optional fields; it does not remove the key. To
restore an unset field, remove its key in the text editor. Ambiguous authored
state can disable the grid while the valid preview remains visible.

Property editing requires the LSP's `userFormPropertyEdit` capability. Older
geometry-capable servers still allow move/resize; preview-only servers remain
read-only. Editing FormSpec source does not require Excel and does not apply
changes to a workbook automatically.

## MultiPage and TabStrip

Keep the authored `controls` list flat and connect a `MultiPage` to its `Page`
controls with `parentId`; put page content controls under each `Page`. A
`TabStrip` is separate: it has its own `tabs` array and does not contain
controls. Each tab has a required `name` and may define `caption`,
`controlTipText`, `tag`, `accelerator`, `enabled`, and `visible`. Omit `tabs`
to leave an existing collection unspecified; use `tabs: []` to author an empty
collection.

`selectedIndex` is zero-based for a nonempty Page or Tab collection. The
normalized no-selection value is `-1`, including when the collection is empty.
Page geometry comes from its owning MultiPage, so do not author Page
`left`/`top`/`width`/`height`. See the [UserForm specification](../reference/userform-spec)
for the complete schema, snapshot shape, and generation details.

The Designer's tab-order editor lists direct controls for one form, Frame, or
Page parent when the LSP metadata exposes `tabIndex`. Hidden and disabled
controls remain in the list; Page controls and TabStrip tab records do not.
Initial order uses authored numeric `tabIndex`, ascending, then z-order for
ties; absent or null indexes follow in z-order. Apply assigns contiguous
zero-based indexes in one edit. If the source changes before Apply, the pending
order is discarded and rebuilt from the document.

## Image pictures

An Image control can load a project-root-relative BMP or JPEG asset in pure-Go
`pack` and `push --backend file`:

```yaml
- id: brand-mark
  name: BrandMark
  type: Image
  picture:
    path: src/forms/assets/brand.jpg
```

Assets are confined to the project after resolving symlinks and junctions, and
are limited to 16 MiB and 16 million pixels. Use `picture: { remove: true }`
to clear an existing picture. Omitting `picture` preserves it. File pull stores
supported embedded images in `src/forms/assets/<full-SHA-256>.bmp|jpg`; those
files are not automatically deleted. File push loads, validates, and
fingerprints referenced assets only; unreferenced assets are retained and
ignored. Excel may normalize JPEG input to BMP when saving. Form build, both
push backends, and pack share the pure-Go compiler for picture validation and
persistence; Excel-backed build reflects that compiled result through the
bridge. See the [UserForm specification](../reference/userform-spec) for
invalid values and full field details.
