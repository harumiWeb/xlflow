# Manage UserForms

Use `form snapshot` to capture Designer state, `form build` to create from a persisted specification, and `form export-image` to verify the rendered result. In sidecar mode, Designer structure lives in `src/forms/specs/` and code-behind in `src/forms/code/`.

```bash
xlflow form snapshot UserForm1 --json
xlflow form build src/forms/specs/UserForm1.yaml --json
xlflow form export-image UserForm1 --json
```

Run `lint` and `form build` preflight before opening Excel. Do not edit generated `.frm` code-behind independently when sidecar mode is authoritative; see the [UserForm specification](../reference/userform-spec).

## Visual layout in VS Code

Open a canonical `.yaml`, `.yml`, or `.json` FormSpec directly under your
configured forms root's `specs` directory, then use **Reopen Editor With... →
xlflow UserForm Designer**. Select a control to drag it or resize with one of
eight handles. Arrow keys move 1 point; Shift+Arrow moves 10 points. Holding
an arrow key creates one edit when the keys are released. Escape cancels an
unfinished interaction. Use VS Code Undo/Redo to restore completed edits.

The grid displays 8-point spacing and has a separate snapping toggle. Snapping
applies to mouse movement and resize; keyboard movement retains its 1/10-point
steps. Zoom ranges from 50% to 200% and does not change source coordinates.
These settings and selection are temporary Designer state.

Movement and resize stay inside the parent's displayed content area. Shrinking
a Frame or MultiPage stops before its children would overflow, including
children on other Pages. Existing out-of-bounds controls are not automatically
corrected; controls larger than their parent must be resized to fit before
moving. Direct source edits remain available for layouts outside these UI
constraints. Form and Page bounds are not directly editable on the canvas.

Visual operations produce local source edits and retain unrelated formatting.
Invalid source retains the last valid preview and disables editing until fixed.
The Toolbox contains Pointer and Label, TextBox, ComboBox, ListBox,
CommandButton, CheckBox, OptionButton, ToggleButton, SpinButton, ScrollBar,
Image, Frame, MultiPage, and TabStrip. Choose a type, then click the canvas to
insert a root control; Escape cancels placement. The Designer uses the same
type-specific default sizes as its preview and clamps the position to fit the
form. A form smaller than the default control size rejects insertion. New
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

Placement inside Frames, reparenting, and Page or Tab collection editing remain
future work tracked by [Issue #920](https://github.com/harumiWeb/xlflow/issues/920) and
[Issue #921](https://github.com/harumiWeb/xlflow/issues/921). Images remain
placeholders.

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
ignored. Excel may normalize JPEG input to BMP when saving. Excel-backed `form build` and apply APIs reject `picture` until
they implement the same contract. See the [UserForm specification](../reference/userform-spec)
for invalid values and full field details.
