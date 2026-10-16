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
eight handles. You can also right-click in the source editor and choose
**xlflow: Open in Designer**; opening keeps unsaved source text intact.
Arrow keys move 1 point; Shift+Arrow moves 10 points. Holding
an arrow key creates one edit when the keys are released. Escape cancels an
unfinished interaction. Use VS Code Undo/Redo or the Canvas toolbar's Undo/Redo
buttons to restore completed edits; both use the same native history.

Designer actions are kept in a shared draft for the FormSpec. The Designer's
**YAML/JSON not applied** badge shows whether that draft has been published to
the source. The VS Code tab's dirty dot tracks its internal checkpoint, so the
dot can clear while the not-applied badge remains.

VS Code Save, Save All, menu Save, and AutoSave durably checkpoint the draft in
xlflow's extension storage; they do not write YAML, YML, or JSON. To publish,
use the Designer Save button or the Designer-scoped Ctrl+S/Cmd+S command. This
validates the current draft and saves it through the latest source document.
Your global AutoSave setting is unchanged, and AutoSave for regular source-text
editing continues normally. Pending drafts survive closing and reopening the
Designer and restarting VS Code. Backup and hot-exit recovery remain supported.

Use Revert to discard the pending Designer draft and adopt the latest YAML/JSON
source buffer. Save As explicitly publishes to the destination you choose. A
corrupt saved draft displays a recovery error and remains available until an
explicit Save or Revert resolves it.

Changes you make in the regular text editor appear in the Designer immediately.
Non-overlapping changes merge with a dirty Designer draft. If the changes
conflict, the Designer retains both versions and disables source publication
until you choose **Use YAML/JSON changes** or **Use Designer draft**. Native
checkpoint saves remain available. A whole-document choice is confirmed
explicitly when the changes cannot be merged safely.

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
insert a control or target a Frame or active Page when its capabilities are
available; Escape cancels placement. The Designer uses the same type-specific
default sizes as its preview and clamps the position to fit the destination. A
destination smaller than the default control size rejects insertion. New
controls have stable `control-<UUID>` IDs and conflict-free generated VBA
names. Insertion selects the new control and returns to Pointer. A new
MultiPage starts with two Pages and `selectedIndex: 0`; its control and both
Pages are one atomic, undoable edit. A new TabStrip starts with `Tab1` and
`Tab2` and `selectedIndex: 0`. Existing authored empty MultiPages and
TabStrips remain unchanged.

Delete the selected control with the toolbar or Delete/Backspace while the
canvas has focus. Backspace also works in Structure. Leaf controls are removed
immediately. When a control has descendants,
a confirmation shows its name and descendant count; confirming removes the
subtree as one undoable edit. Designer deletion does not intercept Delete in a
focused input, so normal character deletion continues to work. Direct deletion
of a Page is disabled. Structural editing requires the LSP's separate
`userFormStructuralEdit` capability. Adding a MultiPage also requires
`userFormPageEdit`; adding a TabStrip also requires `userFormTabEdit`. Other
Toolbox controls retain their existing capability requirements. Older servers
keep their advertised geometry/property support and do not offer unsupported
structural actions. Source edits do not change code-behind. Publish source with
the Designer Save button or its scoped Ctrl+S/Cmd+S command; AutoSave
checkpoints the draft only.

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

### Workspace and navigation

The compact toolbar above the canvas groups alignment, distribution, size, and
z-order commands in menus. Alignment and size appear for at least two sibling
controls; distribution needs three. **More actions** contains Duplicate and
Properties. Grid, Snap, and Zoom are in the footer.

The right pane has **Properties**, **Structure**, and **Tab order** tabs.
Structure lets you select overlapping or invisible controls by name. Selecting
a control on an inactive Page reveals that Page without editing its source
selectedIndex. Expand containers with the disclosure arrows; use arrow keys
for navigation and Enter/Space to select. Structure does not drag controls into
a different parent. Below 900 px, toolbar buttons open Toolbox or the right
pane over the canvas.

**Open source beside Designer** opens the YAML/YML/JSON editor alongside the
Designer. Use the header Save action to publish Designer changes as described
above; opening source or changing a pane does not publish them.

### Open event code

Double-click a standard control or the form background to open its default
event in the VBA editor. For example, TextBox opens Change, CommandButton and
ListBox open Click, and MultiPage opens Change. **Open event…** in the selected
control's actions or the Structure context menu lets you choose another event,
including UserForm Initialize. An existing handler opens at its body; a missing
handler is appended with the VBE-compatible declaration and an empty body.

If the Designer draft or its YAML/YML/JSON source is unsaved, choose **Save and
Continue** or cancel. Only that form's source is saved. Generated code stays
unsaved in its editor and uses normal code-editor Undo independently of layout
Undo. The action does not build a workbook or run Excel.

The code destination follows `[userform].code_source` in `xlflow.toml`:
`sidecar` edits `<forms-root>/code/<FormName>.bas` and creates it when missing;
`frm` edits the VBA portion of the existing matching `.frm`. Missing or
ambiguous frm files require correction before generation. Existing dirty code
is preserved, and existing declarations are never replaced automatically.
Renaming or deleting a control does not rename or remove its old handlers.

Event selection supports the standard Toolbox controls and UserForm itself.
Page and Tab entries use the owning MultiPage or TabStrip. Image opens
BeforeDragOver, matching the tested VBE behavior; its hidden Click/DblClick
members are not offered. Custom ActiveX controls require handwritten handlers.
Update xlflow if the installed LSP does not support event navigation.

`form build` always consumes matching sidecar code, including in frm mode.
When editing frm code, use the project's existing frm push workflow or migrate
explicitly to sidecar mode before a canonical form build.

### Edit properties

Select the background for form properties, a control for its properties, or a
Page tab for Page properties. The Property Grid uses metadata from the running
LSP in name/value columns. Property names remain English regardless of the UI
language; categories, help and actions are localized. Categories can collapse.
Strings/numbers confirm with Enter or on blur; Escape cancels the draft.
Booleans use checkboxes and enums use ComboBoxes. There are no type or `null`
selectors: enter a concrete value directly. Existing unset and `null` fields
stay unchanged until you confirm a concrete value; opening or focusing an input
does not fill defaults. Empty strings, false and zero remain valid values.
Fields supporting any scalar keep their existing concrete type; unset or null
values use text input without inferring types from what you type. FormSpec files
still support `null`; edit the source to set null or remove a property.

Dedicated Font and Color dialogs preview both a sample and the Canvas. Font
supports family, point size, styles and numeric charset, with representative
families and free input. Colors retain RGB, named system or raw OLE identity;
preview font fallback and system colors may differ from Excel. Applicable
built-in fields come from Go metadata, not a bundled frontend schema. Custom
ActiveX and individual Tab entries do not acquire typed editors. Page has no
Font/ForeColor/BackColor/BorderStyle, and CommandButton has no TextAlign.
TextProps font sizes use 0.05 pt increments; empty font names and non-ASCII
updates to an existing StdFont are rejected without rewriting opaque data.

Image controls offer a native BMP/JPEG picker and picture alignment/size modes.
A file inside the project uses its relative path. An external file is imported
only on OK into `<forms-root>/assets/<full-lowercase-SHA256>.bmp|jpg`; identical
content is reused and successful assets remain after Undo for Redo. Cancel
creates no asset. Remove explicitly authors `picture: { remove: true }`.
Position changes only edited axes in one action: Left/Top/Width/Height for one
control, Width/Height for the form; Page and multiple selection are excluded.

Opening or cancelling a modal creates no draft edit or history. OK creates one
native undoable Designer action; explicit Designer Save still publishes source.
If source changes or the LSP restarts, the modal keeps its input, removes its
Canvas preview and disables OK. Cancel and reopen against the refreshed source.
A failed confirmation retains input and its error. Modal input never fills
source fields from preview defaults. Legacy properties migrate only when that
specific property is confirmed, retaining unrelated property-bag entries.

Scalar property editing requires `userFormPropertyEdit`. Typed dialogs require
editor capability version 1; older servers retain their advertised scalar or
geometry behavior. Source syntax that cannot be edited safely can disable the
grid while leaving a valid preview visible. Page geometry, observed fields,
arbitrary property bags and generic property removal remain outside the grid.
Changing `form.name` does not rename files, code sidecars or VBA references.
FormSpec source editing does not require Excel or automatically apply changes
to a workbook.

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

The right pane's **Tab order** editor lists direct controls for one form, Frame, or
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
