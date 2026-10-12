# ADR-0070: UserForm Designer Page and Tab Collection Editing

## Status

Accepted

## Context

ADR-0068 introduced source-preserving root control insertion and subtree
deletion; ADR-0069 added Frame placement and parenting. MultiPage Pages and
TabStrip tabs have different canonical shapes: Pages are `Page` controls in
the flat `controls` array, while TabStrip tabs are records under `tabs`.
Treating either collection as ordinary independent controls would lose its
owner, order, selection, or tab metadata. The existing FormSpec parser and
validator already define the accepted topology and fields.

Issue #921 adds visual Page and TabStrip editing while retaining the
TextDocument as the source of truth and preserving localized YAML/JSON edits.

## Decision

- Keep Pages as flat `Page` controls with `parentId` referencing one MultiPage;
  keep TabStrip entries as records in that control's `tabs` array. Do not add a
  nested persisted Page format or synthetic Page controls for TabStrip tabs.
- Track the Designer's active Page by its stable control ID as transient UI
  state. Selecting a Page in the rendered headers or Page list changes only
  the preview. It does not write `selectedIndex` or any other source field.
  The Page list exposes Pages by ID. Page `name` and `caption` use the existing
  LSP property metadata and source-state contract; Page geometry remains
  derived and unavailable as an authoring property.
- Add, delete, and reorder Pages through dedicated semantic operations.
  Append a Page to its owning MultiPage, confirm deletion when the Page has
  descendants and remove them as one cascade, and expose explicit up/down
  reordering controls. Empty Pages are deleted immediately.
  A Page remains attached to its current MultiPage; it cannot be reparented to
  another MultiPage. Ordinary controls can be inserted into the active Page,
  dropped there, or assigned to a Page through the destination picker. Do not
  use hover-tab drag to reorder Pages.
- Keep TabStrip tabs independent of controls. The tab editor lists every tab,
  including `visible: false` entries, and edits `name`, `caption`, `enabled`,
  `visible`, `tag`, `controlTipText`, and `accelerator` using LSP-provided
  metadata and authored source values. Add appends, remove uses the current
  zero-based index, and reorder uses explicit source and destination indexes.
  TabStrip `selectedIndex` remains a control property handled by the existing
  property metadata path.
- Add the semantic operations `addPage(control)`,
  `removePage(controlId, cascade)`, `reorderPage(controlId, index)`,
  `addTab(controlId, tab)`, `removeTab(controlId, index)`,
  `reorderTab(controlId, index, toIndex)`, and
  `setTabProperty(controlId, index, field, value)`. Page collection operations
  require `capabilities.experimental.userFormPageEdit`; TabStrip collection
  operations require `capabilities.experimental.userFormTabEdit`. These
  capabilities are independent of geometry, general properties, structural
  control edits, and parent edits. Page property edits and MultiPage/TabStrip
  `selectedIndex` edits continue to use `userFormPropertyEdit`. Adding an
  ordinary control to a Page also requires `userFormStructuralEdit` and
  `userFormParentEdit`; changing its Page parent requires `userFormParentEdit`,
  and both pointer dragging and the destination picker require the existing
  geometry capability for `setParent` coordinates.
- Preserve selection by item identity when a collection operation changes its
  order. For an explicitly authored numeric `selectedIndex`, adjust its index
  to keep the same Page or tab selected after reorder or deletion. If the
  selected item is deleted, select index zero when entries remain and `-1`
  when the collection becomes empty. When the first entry is added to an
  explicitly authored `selectedIndex: -1`, set it to zero. Leave an absent or
  explicit-null `selectedIndex` absent or null. This bookkeeping does not
  persist the transient Page selection.
- Preserve observed data as observation. Do not copy observed or normalized
  defaults into authored source. Keep semantic edits local to the affected
  control, Page, or `tabs` sequence in YAML/YML/JSON; apply a batch against the
  original source and validate its final state with canonical FormSpec
  validation. A failed final validation returns no edits and publishes no
  partial source.
- Limit this decision to source editing. It adds no Excel, COM, VBIDE, compiler,
  or runtime behavior.

## Consequences

- Users can author Page topology and TabStrip metadata without changing the
  canonical FormSpec model or editing unrelated source formatting.
- The editor needs separate capability gates and collection-aware UI because
  Page controls and TabStrip records have different identities and ownership.
  Servers that omit either capability continue to expose only their other
  advertised edit families.
- Page and tab selection remains convenient for previewing without silently
  authoring selection state. Explicit numeric selection needs source updates
  when its selected collection member moves or is removed; absent and null
  states remain distinguishable.
- Confirmed Page deletion can remove nested controls. Users receive one
  undoable source edit, and canonical validation remains the final authority.
- This contract does not establish Excel or VBE behavior; runtime collection
  semantics remain governed by the existing bridge and compiler contracts.

## Alternatives Considered

1. **Serialize MultiPage Pages as a nested collection.** Rejected because the
   canonical FormSpec already uses flat controls and `parentId`; a second shape
   would duplicate topology and validation.
2. **Represent TabStrip tabs as Page or child controls.** Rejected because tabs
   are metadata records, have no `parentId`, and must not participate in
   control layout or hierarchy.
3. **Persist the active Page on every click.** Rejected because preview
   navigation is transient and must not create an authored `selectedIndex`.
   Explicit numeric selection is preserved only when a collection mutation
   would otherwise select a different member.
4. **Reuse general control add/remove/reorder for Pages and tabs.** Rejected
   because Page ownership and cascade behavior differ from ordinary controls,
   while tabs are not controls. Dedicated operations make each target and
   collection unambiguous.
5. **Gate these operations on existing structural or parent capabilities.**
   Rejected because older servers can advertise those capabilities while
   lacking Page or Tab collection support. Separate capability signals allow
   independent compatibility negotiation.
6. **Reorder by dragging over tab headers.** Rejected in favor of explicit
   up/down controls, which make the target order deliberate and avoid hover
   gestures changing authored collection order.

## Related

- Issue #921 (parent: Issue #914)
- `docs/adr/ADR-0065-userform-source-preserving-semantic-edits.md`
- `docs/adr/ADR-0067-userform-designer-property-metadata.md`
- `docs/adr/ADR-0068-userform-designer-structural-control-edits.md`
- `docs/adr/ADR-0069-userform-designer-frame-parenting.md`
- `docs/specs/userform-designer.md`
- `docs/specs/userform-semantic-edit.md`
- `docs/specs/userform-multipage-tabstrip.md`
- `internal/vba/userforms/spec/contract.go`
- `internal/vba/userforms/spec/edit/structural.go`
- `internal/lspserver/userform_edit.go`
- `editors/vscode/src/userFormEditor/collections.ts`
- `editors/vscode/webview/userFormDesigner/CollectionEditor.tsx`
- `internal/lspserver/userform_edit_test.go`
- `editors/vscode/test/userFormCollections.test.tsx`
