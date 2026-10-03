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

## Excel bridge and inspection shape

The Designer writer creates Page controls through the MultiPage Pages
collection and creates controls under a Page through that Page's Controls
collection. Before adding authored pages, it removes every current page by
exact page name so the resulting Page collection matches the authored list,
including an empty list. It removes standalone TabStrip tabs by numeric
collection index. These collection-specific removal forms reflect measured
Excel COM behavior: name removal works for Pages, while indexed removal works
for Tabs.

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

## Page geometry and deterministic generation

Page `left`, `top`, `width`, and `height` are derived from the owning
MultiPage, not authoring inputs. Inspection may preserve observed Page bounds
under snapshot-only `observed` data. The compiler rejects Page geometry edits
instead of treating the bounds as independent values.

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
