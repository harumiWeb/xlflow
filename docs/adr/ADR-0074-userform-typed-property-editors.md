# ADR-0074: UserForm Typed Property Editors

## Status

Accepted

## Context

ADR-0067 makes the running Go/LSP contract authoritative for property metadata
and authored presence. Scalar inputs cannot safely express partial fonts,
system colors, or image import intent. ADR-0073 keeps Designer actions in a
shared source-text draft until explicit publication. A property dialog needs
an additional temporary preview without creating draft history on every input.
ADR-0072 requires the same pure-Go compiler across workbook authoring paths.

## Decision

- Supplement ADR-0067 with opt-in editor metadata version 1. The server
  advertises `userFormPropertyEditors: 1`; an opted-in preview request carries
  `propertyEditorVersion: 1`. Go supplies editor kinds, child descriptors,
  applicability and allowed values. The Webview registry selects components,
  labels and layout, never a second validation schema. Older clients retain
  scalar metadata and older servers retain their existing capability behavior.
- Keep FormSpec `schemaVersion: 1`. Model partial `font` objects with point
  sizes and numeric `charset`, and `foreColor` / `backColor` objects whose
  `kind` is `rgb`, `system` or `ole`. Preserve system-color identity and unknown
  raw OLE values rather than replacing them with resolved RGB colors.
- Read legacy property-bag values. Confirmation migrates only the edited
  property's legacy entry into its typed field in the same transaction.
  Equivalent coexisting values are accepted; conflicts are diagnosed.
  Preserve omitted members and unrelated persisted font and opaque data.
- Permit typed values only for explicitly modeled editor properties through
  `setFormProperty` / `setControlProperty`. This does not authorize arbitrary
  object writes or generic property deletion. Use one `setGeometry` operation
  for Position confirmation, with optional changed axes and optional
  `controlId`; its absence targets the form's width and height only.
- Keep modal state and Canvas preview separate from the shared Designer draft.
  Opening, typing and Cancel create no source edits or history. OK validates
  through the LSP and creates one native undoable Designer action; explicit
  Designer Save still publishes source under ADR-0073.
- Bind a modal to its URI, target identity and document version. A source
  change or LSP restart retains the modal draft, clears its temporary Canvas
  preview and disables OK. Reopening obtains current metadata. Do not rebase
  an unconfirmed property draft automatically. Failed commits retain input
  and diagnostics; pending edits block another editor from opening.
- Use current font, representative candidates and free input instead of OS
  font enumeration. Canvas font fallback and system-color rendering are
  approximate previews, not Excel rendering guarantees.
- Limit the picture picker to built-in Image controls and BMP/JPEG. A native
  extension picker stages validated bytes; an external file is imported only
  on OK to `<forms-root>/assets/<full-lowercase-SHA256>.bmp|jpg`. Reuse equal
  content and never overwrite different content. Retain assets after Undo for
  Redo. Only explicit `picture: { remove: true }` removes a picture.
- Preserve the shared canonical compiler boundary for font, color, enum and
  picture persistence. Do not add an imperative per-property COM fallback.

## Consequences

- New editor features can negotiate independently without breaking scalar
  clients. Source-preserving edits must retain comments, precision, authored
  absence/null/empty/false/zero, and untouched typed members.
- Temporary previews improve feedback but require explicit invalidation,
  focus handling and modal lifecycle tests. Native font availability and
  theme colors can differ from the resulting Excel form.
- Asset staging adds file lifecycle work. Cancel must not copy an external
  asset; failed confirmation may clean up only a newly created, unreferenced
  asset. Successful edits retain files for history.
- Font, Color and Position target applicable built-in controls and the form;
  Page appearance properties, CommandButton TextAlign, individual Tab records,
  custom ActiveX property editors, OS font catalogs,
  PNG/GIF and non-Image picture properties remain outside scope.
- Pure source tests do not establish Excel compatibility. Persisted Font,
  Color and Image behavior requires focused Windows/Excel release validation.

## Alternatives Considered

1. **Frontend-owned applicability and schema.** Rejected because installed
   extension and server versions can differ.
2. **Resolve every color to RGB.** Rejected because system colors and raw OLE
   values carry meaning beyond the current theme.
3. **Apply each dialog input immediately.** Rejected because Cancel would need
   to undo authored changes and history would contain preview-only actions.
4. **Import pictures when selected.** Rejected because Cancel would leave new
   project assets despite creating no authored change.
5. **Replace entire font records or migrate every legacy property.** Rejected
   because a narrow edit must preserve unrelated and unknown persisted state.

## Related

- `docs/adr/ADR-0067-userform-designer-property-metadata.md`
- `docs/adr/ADR-0072-userform-canonical-compiler-and-persisted-dimensions.md`
- `docs/adr/ADR-0073-userform-designer-deferred-save-document-lifecycle.md`
- `docs/specs/userform-property-editors.md`
- `internal/vba/userforms/spec/appearance.go`
- `internal/vba/userforms/spec/edit/property_editors.go`
- `internal/vba/userforms/oforms/font.go`
- `editors/vscode/src/userFormEditor/protocol.ts`

## Supersedes

- The scalar-only editor scope in ADR-0067, for negotiated editor version 1.
  Its Go metadata authority and authored-state rules remain in effect.

## Superseded by

- None
