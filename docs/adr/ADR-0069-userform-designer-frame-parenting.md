# ADR-0069: UserForm Designer Frame Parenting

## Status

Accepted

## Context

Issue #920 supplements ADR-0068's root-only Toolbox with Frame placement and
reparenting. The existing renderer already nests controls through canonical
`parentId`, but geometry gestures clamp movement to the current parent. Older
servers advertise structural edits without implementing parent edits. YAML's
host-neutral `setParent` already exists; JSON needs equivalent source edits.

## Decision

- A normal pointer drag explicitly commits a parent change when dropped into
  a Frame or onto the form background. Highlight the destination and preview
  the converted position before committing. Geometric overlap during loading,
  text editing, resizing or keyboard movement never creates a relationship.
- Keep origin conversion, content insets, stacking and target hit testing in
  the shared TypeScript geometry layer. Render an in-progress drag in form
  coordinates so the original parent's clipping cannot hide it. Moving a Frame
  carries its descendants without rewriting their local coordinates.
- Preserve the pointer-to-control offset when changing coordinate spaces.
  Apply snapping before boundary clamping. A control that fits the destination
  may be clamped to its edge; one larger than its destination cannot be dropped.
  An invalid destination never silently falls back to a different parent.
- Use one `setParent` operation carrying `parentId`, `left` and `top` for a
  parent change, and one native undoable edit. Root parenting is an empty
  `parentId`. Same-parent moves retain `moveControl` and existing geometry rules.
- Add an independent `userFormParentEdit` capability. Parent-changing drags
  also require geometry support; parented additions also require structural
  support. Servers without this capability retain root placement and existing
  geometry/property behavior.
- Limit Designer parent changes to built-in non-Page controls whose current
  direct parent is a Frame or the form. New direct parents are Frames or the
  form. Direct Page relationships and Page/Tab authoring remain Issue #921.
  The extension host and LSP enforce scope; canonical FormSpec validation
  remains authoritative for relationship validity. The generic edit API is not
  limited to Designer scope.
- Extend JSON source-preserving `setParent`, including relocating a legacy
  nested source item to the canonical top-level sequence when its parent
  changes. Retain its subtree, effective ordering and unrelated source.
- Show a clickable selection breadcrumb instead of adding a structure tree.
  Selecting an ancestor changes transient UI selection only.

## Consequences

- Users can create and rearrange nested Frames without authoring `parentId`
  manually, with previews and native Undo/Redo.
- Edge clamping can slightly change the apparent position. Previewing the
  exact final placement makes that correction visible before the drop.
- Parenting compatibility is explicit rather than inferred from an older
  structural capability. Page-aware editing remains a separate feature.
- Source changes establish no Excel/COM/VBE runtime behavior.

## Alternatives Considered

1. Require Alt-drag to change parents. Rejected in favor of normal dragging
   with visible destination feedback, as selected for Issue #920.
2. Reject every partially overflowing drop. Rejected because previewed edge
   clamping is more useful when the control fits inside the destination.
3. Add a full structure tree. Deferred; clickable ancestors provide the
   required nested selection without another editor panel.
4. Reuse `userFormStructuralEdit`. Rejected because older servers advertise
   it while explicitly rejecting parented addition and `setParent`.

## Related

- Issue #920 (parent: Issue #914); Issue #921
- `docs/adr/ADR-0068-userform-designer-structural-control-edits.md`
- `docs/specs/userform-designer.md`
- `docs/specs/userform-semantic-edit.md`
- `editors/vscode/src/userFormEditor/geometry.ts`
- `editors/vscode/webview/userFormDesigner/DesignerCanvas.tsx`
- `editors/vscode/test/userFormParenting.test.tsx`
- `internal/lspserver/userform_edit.go`
- `internal/vba/userforms/spec/edit/json_source.go`
