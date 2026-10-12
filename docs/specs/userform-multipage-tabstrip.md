# UserForm MultiPage and TabStrip Contract

This document records the v1 UserForm schema and backend rules for MultiPage,
Page, and TabStrip controls. It follows the canonical Go contract in
`internal/vba/userforms/spec` and the compiler and OFORMS edit APIs. It does not
change the `schemaVersion: 1` discriminator or introduce backend fallback.

## Authored control model

`FormSpec.controls` remains a flat array connected with `parentId`:

```text
MultiPage -> Page -> child controls
TabStrip  -> independent tabs array; no child controls
```

- A MultiPage may contain only Page controls.
- A Page must have a MultiPage parent and can contain ordinary controls.
- A TabStrip cannot contain controls. Its tabs are not FormSpec controls and
  have no `parentId`.
- `MultiPage.selectedIndex` and `TabStrip.selectedIndex` are zero-based
  indexes for their respective collections. Nonempty collections require
  `0..count-1`; empty collections require logical selection `-1`.

The canonical Go `FormSpecTab` JSON/YAML fields are `name` (required),
`caption`, `controlTipText`, `tag`, `accelerator` (optional strings), and
`enabled`, `visible` (optional booleans). Tab names are unique within their
owning TabStrip, case-insensitively. These fields are scoped to each tab
record; they are not general TabStrip control properties.

For edits to an existing form, an omitted `tabs` field leaves the collection
unspecified and preserves the existing collection. An explicit `tabs: []`
authors an empty collection. When `tabs` is authored, its entries and order
define the requested collection; omitted optional metadata on retained tabs
can be preserved by matching the tab name. MultiPage pages are authored as
Page controls in the flat control list, not as a TabStrip `tabs` array.

## Designer collection editing

The VS Code Designer retains these canonical shapes. Its active Page is
transient UI state keyed by the Page control's stable `id`; clicking a rendered
Page header or Page-list entry changes the preview only and does not write
`selectedIndex`. The Page list identifies each Page by control ID. Page
`name` and `caption` use the existing LSP property metadata and authored-value
contract. Page `left`, `top`, `width`, and `height` remain derived and are not
editable properties.

Page collection operations are `addPage(control)`,
`removePage(controlId, cascade)`, and `reorderPage(controlId, index)`. Adding
appends a Page to its referenced MultiPage. Deleting a Page requires explicit
Designer confirmation and removes its descendants with `cascade: true` as one
source edit. Reordering uses explicit up/down actions within the same
MultiPage; it does not move the Page to another MultiPage or use hover-tab
dragging. Ordinary controls can be inserted into the active Page, dropped on a
Page, or assigned to a Page through the destination picker. Those controls may
change Page parent; Page controls themselves cannot be reparented.

TabStrip editing keeps tabs as independent records and presents every record,
including entries with `visible: false`. The editor uses LSP metadata and
authored source values for `name`, `caption`, `enabled`, `visible`, `tag`,
`controlTipText`, and `accelerator`. The operations are `addTab(controlId,
tab)` (append), `removeTab(controlId, index)`,
`reorderTab(controlId, index, toIndex)`, and
`setTabProperty(controlId, index, field, value)`. Tab records never appear as
Page or child controls. TabStrip control-level `selectedIndex` remains an
ordinary property edit through the existing property metadata path.

The LSP advertises Page operations with the independent
`capabilities.experimental.userFormPageEdit` capability and TabStrip
collection/property operations with
`capabilities.experimental.userFormTabEdit`. Page `name`/`caption` and
MultiPage or TabStrip control-level `selectedIndex` use
`userFormPropertyEdit`. Insertion into a Page also needs the existing
structural and parent capabilities. Changing an ordinary control's Page parent
needs the existing parent capability and geometry capability for a pointer
drag; the destination picker does not require geometry. These gates are
independent so clients keep each edit family a server advertises.

Collection operations preserve the selected member by identity when an
explicit numeric `selectedIndex` is authored. Reordering updates the numeric
index to keep that same Page or tab selected. Deleting the selected member
falls back to index zero if members remain, or logical `-1` when the collection
is empty. Adding the first member changes an explicitly authored `-1` to zero.
An absent or explicit-null `selectedIndex` remains absent or null; preview-only
Page selection never authors the field. Observed values and normalized
defaults are not copied into authored source.

Page and Tab operations make localized YAML/YML/JSON source edits. They retain
unrelated comments, whitespace, and ordering, and leave observed data
untouched. A transaction is validated against the canonical FormSpec after
all requested operations have been applied. A failed target resolution, source
edit, or final validation returns no edit and publishes no partial result.
These Designer operations change FormSpec source only and add no Excel, COM,
VBIDE, compiler, or runtime semantics.

## Excel bridge and inspection shape

The Designer writer creates Page controls through the MultiPage Pages
collection and creates controls under a Page through that Page's Controls
collection. Before adding authored pages, it removes every current page by
exact page name so the resulting Page collection matches the authored list,
including an empty list. It removes standalone TabStrip tabs by numeric
collection index. These collection-specific removal forms reflect measured
Excel COM behavior: name removal works for Pages, while indexed removal works
for Tabs.

The writer creates all siblings before assigning their TabIndex values in
ascending order, then verifies the final sibling indexes. MSForms can clamp
an index during creation and renumber previously created siblings. List items
are populated before writing selection Text/Value. Omitted nested TabStrip
state is captured with the inspection type resolver, including containers
whose ProgId is unavailable. Caption normalization doubles embedded quotes
and leaves multiline captions in the exported FRX without adding source lines.

Designer and runtime inspection represent MultiPage Pages as child controls
and expose TabStrip tabs as independent tab metadata. TabStrip is not treated
as a child-control container. Inspection reports the logical selection as
`selected_index`; authored FormSpec uses `selectedIndex`. The internal hidden
TabStrip used by a MultiPage is not emitted as a second user control.

For a nonempty collection, Excel's `Value` supplies the selected index for both
MultiPage and TabStrip. An empty collection is normalized to logical
`selected_index: -1`. Excel can retain a stale raw MultiPage `Value` after its
last Page is removed; the bridge does not read or assign `Value` while that
Page collection is empty. TabStrip reports `-1` naturally when empty. This
normalization describes the public snapshot and does not rewrite raw saved
state.

Saved empty MultiPages may also retain hidden tab arrays. The x stream owns
the empty Page topology, so projection omits those cached entries and no-op
compilation preserves them. Direct TabEdit calls on an empty MultiPage reject
changed cached tab state. Adding a Page through topology replaces the cache
with the final authored Page collection.

After a Page collection changes, topology reconciles the selected Page identity
and synchronizes every remaining Page's Site visible bit to the final selection.
This Site bit identifies the active Page and is separate from the authored
Page `visible` flag stored in the hidden TabStrip. Synchronization also applies
when deleting the selected Page falls back to index zero or an added Page has
an initially visible Site. No-op replay retains existing saved bytes.

Copied observed selection describes the original collection. Compilation
discards that unchanged observation before validating the final topology;
explicit authored selection is checked against the final collection, while
omitted selection follows retained identity. The property-only stage uses
the original collection and its original observed selection.

## Page geometry and deterministic generation

Page `left`, `top`, `width`, and `height` are derived from the owning
MultiPage, not authoring inputs. Inspection may preserve observed Page bounds
under snapshot-only `observed` data. The compiler rejects Page geometry edits
instead of treating the bounds as independent values.
The Designer writer also skips observed Page bounds during snapshot rebuilds;
ordinary controls under a Page still receive their authored geometry.

New pure-Go generation computes standard top-tab Page bounds at fixed 96 DPI
and writes an explicit Tahoma 8.25pt font. The generated geometry does not
query the machine's display DPI or installed font settings.

Existing standard layouts retain their measured Page insets when resized or
when a Page is added/moved into that owner. Mutations requiring unmodeled
nonstandard tab-layout bookkeeping fail before publication. Page metadata may
use its typed fields or known property-bag aliases; conflicting explicit aliases
are rejected before Excel component replacement or pure-Go publication.

## Atomic modeled edits

The `compiler.CompileNew`, `compiler.CompileTemplate`, and
`compiler.CompileEdits` APIs validate the requested model and return a newly
parsed result. The `oforms.ApplyEdits` and `oforms.ApplyTabEdits` mutation APIs
apply edits to an independent clone and return a new model only after encoding
and read-back succeed. On an invalid, stale, or unsupported edit, the call
returns an error without mutating its input model or returning a partially
edited result. Callers publish only a fully validated result; there is no
automatic switch to another backend after an error.
