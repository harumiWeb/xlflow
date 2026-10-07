# ADR-0068: UserForm Designer Structural Control Edits

## Status

Accepted

## Context

ADR-0064 makes the VS Code `TextDocument` the source of truth for the UserForm
Designer. ADR-0065 defines source-preserving semantic edits, ADR-0066 adds JSON
geometry transactions, and ADR-0067 separates property metadata and capability
negotiation. Issue #919 adds the minimum structural operations needed to create
and remove controls visually. Placement must preserve those boundaries while
remaining useful before container and Page-specific editing arrive in Issues
#920 and #921.

## Decision

- Add a Toolbox with Pointer and the 14 built-in control types other than Page.
  Selecting a type and clicking the canvas inserts a root control. Use the
  existing shared type-size defaults, clamp the click position so the control
  fits the form, and reject insertion when the form is smaller than that
  default size. Escape cancels placement; a successful insertion selects the
  new control and returns the Toolbox to Pointer.
- Give each inserted control a `control-<UUID>` ID that remains stable through
  rename. Generate the familiar type-prefixed VBA name using the smallest
  positive suffix free under case-insensitive comparison with all control
  names and the form name. Before dispatch, the extension host rechecks the ID
  and checks the name case-insensitively against the current form and control
  names. These are the Designer's name-conflict guards; Go `Apply` and
  canonical FormSpec validation add no global control-name uniqueness rule.
- Start MultiPage without Page child controls and with `selectedIndex: -1`.
  Start TabStrip with `tabs: []` and `selectedIndex: -1`. Page insertion,
  deletion, tabs editing, placement inside a Frame, and reparenting remain
  outside this decision and are handled by Issues #921 and #920 respectively.
- Extend source-preserving Go semantic edits to JSON `addControl` and
  `removeControl`, including cascade removal through legacy nested source.
  Preserve unrelated source and validate the final result only through the
  canonical FormSpec parser and validators. Do not add global control-name
  uniqueness validation.
- Offer deletion through a toolbar action or Delete while the canvas is
  focused. Do not handle Delete in text inputs. Remove a leaf control
  immediately; when descendants exist, confirm with a host modal that
  identifies the selected control and descendant count, then remove the
  selected control and descendants with `cascade: true` in one `WorkspaceEdit`.
  Disable direct deletion of a selected Page.
- Normalize unsnapped TypeScript placement coordinates with the existing
  `stableDelta` helper to six decimal places before clamping. When snapping is
  enabled, apply the existing grid snap before the form-bound clamp.
- Advertise structural add/remove through the separate
  `capabilities.experimental.userFormStructuralEdit` capability. Keep
  geometry and property capability negotiation independent. The host guards
  queued edits and pending confirmations against changes to unsaved text,
  document version, project and LSP context. Native VS Code history owns
  Undo/Redo; structural edits do not change code-behind or save automatically.

## Consequences

- Users can create a flat, root-level form layout and remove a confirmed
  subtree without losing the original source formatting around the edit.
- A generated name is conflict-free at allocation time, but FormSpec source
  validation continues to accept duplicate names where it already does. Any
  stricter writer-specific constraint remains downstream.
- Users cannot create children inside Frames, reparent controls, or manage
  Pages and TabStrip tabs through this Toolbox. Those workflows remain
  separate scope in Issues #920 and #921.
- Older servers that omit the structural capability do not provide Toolbox
  insertion or deletion. Their existing preview, geometry, and property
  behavior remains governed by the capabilities they advertise.
- The feature changes source documents only. It establishes no COM, VBIDE,
  workbook persistence, or VBE behavior.

## Alternatives Considered

1. **Offer drag-to-create and container targeting.** Deferred because click-to-
   place is sufficient for this scope, while container placement and parenting
   require the hierarchy and coordinate behavior assigned to Issue #920.
2. **Expose Page as a general Toolbox item.** Rejected because a Page requires
   MultiPage-specific ownership and ordering, which belongs to Issue #921.
3. **Skip confirmation for every deletion or remove only the selected node.**
   Rejected for controls with descendants because either choice can silently
   destroy or orphan them. A named confirmation and one cascade transaction
   make the affected subtree explicit and undoable; leaf controls have no
   descendants and are removed immediately.
4. **Add a global duplicate-name validator.** Rejected because canonical
   FormSpec validation does not currently impose that rule. The Toolbox
   allocator can avoid generated conflicts without changing source validity.
5. **Gate structural operations on the geometry or property capability.**
   Rejected because older servers may implement one edit family without the
   other; structural support has its own explicit compatibility signal.

## Related

- Issue #919 (parent: Issue #914)
- Issue #920 (container parenting and hierarchy)
- Issue #921 (MultiPage, Page, and TabStrip editing)
- `docs/adr/ADR-0064-userform-designer-text-document-boundary.md`
- `docs/adr/ADR-0065-userform-source-preserving-semantic-edits.md`
- `docs/adr/ADR-0066-userform-designer-geometry-transactions.md`
- `docs/adr/ADR-0067-userform-designer-property-metadata.md`
- `docs/specs/userform-designer.md`
- `docs/specs/userform-semantic-edit.md`
- `editors/vscode/src/userFormEditor/toolbox.ts`
- `editors/vscode/src/userFormEditor/provider.ts`
- `editors/vscode/src/client.ts`
- `internal/lspserver/server.go`
- `internal/lspserver/userform_edit.go`
- `internal/vba/userforms/spec/edit/edit.go`
- `internal/vba/userforms/spec/edit/json_source.go`
- `editors/vscode/test/userFormStructure.test.tsx`
- `editors/vscode/test/suite/userFormWebview.test.ts`
- `editors/vscode/test/suite/userFormProvider.test.ts`
- `internal/lspserver/userform_edit_test.go`
- `internal/vba/userforms/spec/edit/edit_test.go`
