# ADR-0067: UserForm Designer Property Metadata and Authored States

## Status

Accepted

## Context

ADR-0064 keeps the VS Code TextDocument authoritative; ADR-0065 provides
source-preserving YAML semantic edits; ADR-0066 adds JSON geometry transactions.
Issue #918 adds a Property Grid. A normalized preview can contain observed or
display defaults and cannot establish whether a field was authored. A static
frontend property schema can also disagree with the running LSP version.

## Decision

- Obtain property descriptors from the running project's LSP through
  `xlflow/userFormPreview.propertyGrid`. Derive field types, required status,
  nullability and allowed values from Go FormSpec contracts intersected with
  the semantic editor's supported scalar fields. Keep labels, categories and
  display order in the UI; do not duplicate canonical validation there.
- Return authored field presence and values from original source syntax,
  separately from the normalized rendering document. Distinguish absent,
  explicit null, empty string, false and zero. Never write observed values or
  renderer defaults back as authored input. If authoring state is ambiguous,
  return `propertyError` while retaining a valid preview.
- Supplement ADR-0065 and ADR-0066 with source-preserving scalar property
  replacement and insertion for YAML, YML and JSON. Use `setFormProperty` and
  `setControlProperty` through the existing semantic edit API; validate the
  resulting source with the canonical parser. Reject unsafe or ambiguous
  targeting atomically rather than reserializing the document.
- Require explicit value presence. Permit explicit null only for an optional
  scalar field accepted by canonical validation; null is a source value, not
  property deletion. Property removal remains unsupported.
- Advertise `capabilities.experimental.userFormPropertyEdit: true` separately
  from geometry editing. One property confirmation is one operation and one
  WorkspaceEdit. Retain the existing same-control move/resize pair, URI queue,
  stale-result rejection and native undo/redo. Older servers keep their
  advertised preview/geometry behavior without property editing.
- Changing `form.name` changes only that FormSpec field. File names, code
  sidecars and VBA references are not renamed implicitly.

## Consequences

- The running server owns field applicability and validation, avoiding schema
  drift between extension and LSP versions. Missing capability or metadata
  disables property editing rather than prompting the UI to guess.
- Source-state extraction needs a separate syntax representation even when
  canonical parsing succeeds. Rendering can remain available when safe
  authoring metadata cannot be produced.
- Local JSON edits now cover scalar properties as well as geometry; structural
  JSON edits, property bags, collections and pictures remain outside #918.
- Explicit null requires presence-aware payload decoding and encoding. Users
  must use the text editor to delete a field and restore absence.
- This changes pure source editing, not COM, VBIDE or VBE semantics. Verification
  uses Go/LSP tests and real VS Code history/synchronization checks; Excel is
  not required for this contract and no workbook/runtime claim is made.

## Alternatives Considered

1. **Bundle a TypeScript validation schema.** Rejected because the extension
   may run against a different server and would duplicate canonical rules.
2. **Populate the grid from rendered values.** Rejected because normalization
   and display defaults lose authored presence and can turn observation into
   input without user intent.
3. **Treat null as deletion.** Rejected because null and absence are distinct
   source states and property removal needs its own semantic contract.
4. **Rewrite JSON or limit the grid to YAML.** Rejected because local scalar
   edits can preserve the user's format without whole-document churn.
5. **Reuse only the geometry capability.** Rejected because older geometry
   servers do not implement property operations or authored-state metadata.

## Related

- Issue #918 (parent: Issue #914)
- Structural operation capability negotiation is supplemented by `docs/adr/ADR-0068-userform-designer-structural-control-edits.md`.
- `docs/adr/ADR-0064-userform-designer-text-document-boundary.md`
- `docs/adr/ADR-0065-userform-source-preserving-semantic-edits.md`
- `docs/adr/ADR-0066-userform-designer-geometry-transactions.md`
- `docs/specs/userform-designer.md`
- `docs/specs/userform-semantic-edit.md`
- `internal/vba/userforms/spec/edit/properties.go`
- `internal/lspserver/userform_preview.go`
- `editors/vscode/test/suite/userFormEdits.test.ts`
