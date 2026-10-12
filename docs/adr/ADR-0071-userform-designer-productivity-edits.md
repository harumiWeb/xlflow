# ADR-0071: UserForm Designer Productivity Edits

## Status

Accepted

## Context

Issues #915–#921 established the UserForm Designer's TextDocument authority,
source-preserving semantic edits, geometry, properties, structural editing,
Frame parenting, and Page/TabStrip editing. Issue #922 adds multi-selection,
group layout commands, z-order, copy/paste, and tab-order editing. These actions
must preserve the existing YAML/YML/JSON source contract and remain usable
across Designer documents without turning the Webview into a second source
editor.

## Decision

- Limit group selection to ordinary built-in controls that are immediate
  siblings under one parent: the form root, one Frame, or one Page.
  Ctrl/Cmd-click and Shift-click toggle eligible siblings. Marquee selection
  includes only visible, ordinary built-in immediate children fully inside the
  marquee. Custom controls remain selectable alone but cannot join a group.
  Group movement translates every selected control equally and keeps its
  parent and relative geometry; canvas dragging never resizes a group.
- Use the last selected control as the explicit alignment or size anchor.
  Provide six edge/center alignment commands, horizontal/vertical distribution
  with fixed endpoints and equal edge gaps for three or more controls, and
  same-width, same-height, and same-size commands. Provide the four standard
  z-order commands while retaining the selected controls' relative order and
  leaving `tabIndex` unchanged.
- Treat each productivity action as one semantic operation and one native
  VS Code `WorkspaceEdit`/Undo step. Copy, Duplicate, and Cut use a preparatory,
  read-only copy request. Cut first captures a valid source snapshot, then, in
  the same per-URI edit queue, asks one confirmation with the aggregate
  descendant count across selected roots when descendants exist; leaf-only Cut
  needs no descendant confirmation. If confirmed, submit one atomic
  `removeControls` cascade edit. The shared clipboard receives the captured
  snapshot only after deletion succeeds.
  Cancellation, stale source, failed capture or deletion, or supersession
  before deletion leaves the source and previous clipboard unchanged. Keep all
  existing `userFormEdit` operation-family and batching restrictions unchanged.
- Add `xlflow/userFormCopy` to extract selected source records from the supplied
  unsaved URI/version/text. The extension host retains a typed
  `ClipboardSnapshot` in memory and shares it across Designer instances.
  Preserve source-authored field presence and values, including observations,
  absence, null, empty string, false, and zero; do not synthesize normalized or
  preview defaults. Pasting recursively allocates new IDs and case-insensitively
  unique names across the whole form and rewrites internal IDs and parent
  references. A copied MultiPage includes its Page subtree; a Page cannot be a
  direct copy root. Reject unsupported custom subtrees atomically. Keep
  project-relative picture paths unchanged within the same project and reject
  them across projects. Do not copy image files, event handlers, or code-behind.
- Offset each repeated paste by 8 pt on both axes. Clamp the complete pasted
  group by one translation so it fits the destination; reject it if the group
  cannot fit. Never resize pasted controls to make them fit.
- Provide a parent-local tab-order editor for controls whose LSP metadata
  exposes editable `tabIndex`. Include hidden and disabled controls; exclude
  Page controls and TabStrip tab records. Initialize from authored numeric
  indexes ascending, with ties by z-order; place absent or null indexes last
  in z-order. Apply one complete ordered ID list as zero-based indexes in one
  operation. Discard an uncommitted local order when its source document
  changes.
- Add `capabilities.experimental.userFormProductivityEdit` independently of
  existing capabilities. Productivity actions also require the applicable
  existing edit family: geometry for movement/layout/z-order, property editing
  for tab order, structural editing for cut/delete/paste, and parent/Page
  capabilities when a paste creates those relationships. A missing capability
  disables only the actions that require it.
- Apply the Issue #922 productivity edits locally to YAML/YML/JSON, validate
  the final document canonically, and return no partial edits on failure. The
  existing single-control `reorderControl` operation gains a JSON writer path
  to support atomic group z-order expansion. These decisions change
  no Excel, COM, VBIDE, VBE, compiler, or runtime behavior.

## Consequences

- Multi-selection and layout commands preserve sibling ownership and make each
  user action a reviewable, single-step source edit. Anchor-based sizing and
  fixed-endpoint distribution are deterministic.
- Clipboard contents remain within the extension host and can move between
  Designer documents. A cross-project paste containing authored image paths
  fails safely because the referenced asset is not copied. Unsupported custom
  hierarchies fail as a whole instead of producing a partial subtree.
- Tab order follows the FormSpec's parent boundaries and LSP field metadata;
  visibility and enabled state do not silently remove controls from the order.
- Older servers continue to provide their independently advertised existing
  edit families. They do not gain productivity actions merely by advertising
  geometry, property, or structural support.
- Source-only validation and native editor Undo do not establish workbook or
  VBE behavior.

## Alternatives Considered

1. **Allow mixed-parent selections.** Rejected because a single translation,
   coordinate space, alignment anchor, and parent-local tab-order list would
   become ambiguous.
2. **Resize a group by dragging a selection handle.** Rejected because it
   changes each control's dimensions and relative geometry in one gesture;
   explicit anchor-based size commands provide a predictable result.
3. **Copy through the operating-system clipboard.** Rejected because source
   fragments need a format tag and exact authored-state preservation across
   Designer documents. An extension-host snapshot keeps that contract explicit.
4. **Delete before obtaining a valid source snapshot during Cut.** Rejected
   because a failed snapshot request could otherwise remove source controls
   without preserving their source data for the clipboard.
5. **Sort tab order by visibility or enabled state.** Rejected because those
   properties do not define the authored parent-local order and can change
   independently of `tabIndex`.

## Related

- Issue #922 (parent: Issue #914)
- `docs/adr/ADR-0065-userform-source-preserving-semantic-edits.md`
- `docs/adr/ADR-0066-userform-designer-geometry-transactions.md`
- `docs/adr/ADR-0067-userform-designer-property-metadata.md`
- `docs/adr/ADR-0068-userform-designer-structural-control-edits.md`
- `docs/adr/ADR-0069-userform-designer-frame-parenting.md`
- `docs/adr/ADR-0070-userform-designer-page-tab-collection-editing.md`
- `docs/specs/userform-designer.md`
- `docs/specs/userform-semantic-edit.md`
- `editors/vscode/src/userFormEditor/provider.ts`
- `internal/lspserver/userform_edit.go`
