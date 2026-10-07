# ADR-0065: Source-Preserving UserForm FormSpec Semantic Edits

## Status

Accepted

## Context

Issue #915 established a read-only Designer preview whose VS Code
`TextDocument` remains canonical. Issue #916 proposes a host-neutral Go layer
that can later apply Designer operations without making the Webview or another
host a second source of truth.

The existing FormSpec parser is also a normalizer. Its semantic model can
contain derived observed values, model-only generated control IDs, normalized
z-indices, and a flattened control representation. Serializing that model
after a visual edit can therefore add or reorder authored data unrelated to
the requested change. Whole-document serialization would undermine readable
source review and would discard source formatting and comments.

ADR-0064 keeps the current Designer read-only and requires any future editing
path to update the canonical text document. This decision specifies the
host-neutral semantic editing core only; it does not change that editor
boundary or introduce an LSP or Webview protocol.

## Decision

Supplement ADR-0064 with a source-preserving FormSpec edit API in
`internal/vba/userforms/spec/edit`:

```go
func Apply(input spec.SpecInput, source []byte, operations []Operation) (Result, error)
```

- The original YAML source is canonical. The existing
  `spec.ParseFormSpec` parser and FormSpec validators remain the sole semantic
  and validation authority. The API accepts YAML/YML only, even though the
  shared parser also reads JSON.
- Source validity and writer capability are distinct. `spec.ParseFormSpec`
  rejects a source control without an explicit `id` with `UFV004` before
  editing. `spec.NormalizeFormSpec` may synthesize an ID in an in-memory model,
  but the edit API neither inserts missing source IDs nor addresses
  normalized-only IDs. Canonical source validation currently does not reject
  duplicate control names, so the edit API adds no such rule. A source-valid
  document may still fail stricter downstream compiler or generator checks;
  those checks remain authoritative for writer capability.
- Keep the normalized `spec.FormSpec` used for semantic operations separate
  from the source syntax representation used to find and emit byte edits. Do
  not serialize the normalized model back into the original file.
- Use a serializable tagged `Operation` struct whose `Type` is `OperationType`,
  with `ControlID`,
  `Field`, `Value`, `Left`, `Top`, `Width`, `Height`, `Control`, `ParentID`,
  `Index`, and `Cascade` fields. Support `setFormProperty`,
  `setControlProperty`, `moveControl`, `resizeControl`, `addControl`,
  `removeControl`, `setParent`, and `reorderControl`.
- Require both coordinates for `moveControl` and both dimensions for
  `resizeControl`. Geometry scalars are also valid `setControlProperty`
  fields. `progId` is not an authorable generic property. `addControl` reads
  its optional parent from `Control.ParentID`; a non-empty operation-level
  `ParentID` is rejected as irrelevant input.
- Restrict ordinary property edits to supported scalar fields. Keep IDs and
  types immutable. Observed state, property bags, collections, pictures,
  property removal, and recursive control creation are unsupported. Form
  property paths are explicit and relative to `form`, including `build.*`;
  never silently map one field to another.
- Apply operations sequentially to a private semantic model, then validate the
  final state with the canonical parser and validators. Any failure returns no
  result or partial source.
- Emit deterministic, localized `SourceEdit` ranges in original UTF-8 byte
  offsets. Preserve unrelated source bytes, comments, and spacing. Regenerate
  only the smallest local flow mapping or sequence when needed. Unsafe targeted
  anchors, aliases, or merge-key expansion fail explicitly; there is no
  whole-document rewrite fallback.
- A semantic no-op returns the original bytes with no edits. Operations target
  explicit IDs from source; omitted IDs fail parsing with `UFV004` before
  operations and are never materialized or targeted by the edit API.
- Removal defaults to `cascade: false`; an absent or false value rejects
  removal while descendants remain, and true removes descendants. Reparenting
  retains local coordinates unless new parent-relative coordinates are
  supplied. Reordering uses stable sibling `zIndex` order and does not rewrite
  `tabIndex`.
- Return structured operation errors with operation identity and preserve
  canonical `UFVxxx` validation codes, messages, and suggestions.
- Keep the API independent of VS Code, Webview, LSP, CLI, Excel, COM, and VBE.

The concrete API and behavioral details, including implementation test
requirements, are specified in
`docs/specs/userform-semantic-edit.md`.

## Consequences

- A visual host can request semantic changes while preserving reviewable YAML
  source and unrelated comments and formatting.
- Canonical validation remains centralized, reducing the risk that a host
  accepts a hierarchy or property value rejected by other FormSpec consumers.
- A syntax tree and canonical model must be kept aligned during a transaction.
  This is more work than marshaling the final model, and some flow-style or
  alias-dependent edits must fail when a safe local edit cannot be proven.
- Unsupported collections and property removal limit the first API surface;
  these operations require later explicit contracts.
- The canonical parser's source-valid boundary does not guarantee that a
  downstream writer can consume every parsed document. In particular,
  duplicate control names are not a source-validation error; writer-specific
  constraints stay with the compiler or generator.
- This decision does not make the current Designer editable and does not
  change Excel/VBE persistence behavior. No actual Excel validation is
  required for this pure-Go source contract; its implementation is verified
  with Go tests defined by the specification.

## Alternatives Considered

1. **Marshal the normalized FormSpec as the edited document.** Rejected
   because normalization may synthesize observed fields, model-only IDs,
   z-indices, or a different control shape, causing unrelated source churn and
   comment loss.
2. **Edit only YAML syntax and duplicate FormSpec rules in the syntax layer.**
   Rejected because syntax-level checks would become a second authority for
   control identity, properties, and parent topology. Syntax locates source;
   the canonical model decides semantics.
3. **Let the Webview serialize edits or maintain a parallel form model.**
   Rejected because ADR-0064 keeps the TextDocument as the sole canonical
   source. A future host must call this API through an explicit adapter.
4. **Fall back to complete YAML regeneration when a local edit is difficult.**
   Rejected because an operation could then rewrite comments and unrelated
   formatting without the user's intent. Unsafe targeted syntax fails
   atomically instead.
5. **Support JSON and all FormSpec fields from the start.** Rejected for this
   API because the initial editing contract is localized YAML mutation over
   scalar fields. JSON and collection/property-bag edits need separate source
   preservation and operation contracts.

## Related

- Scalar property editing across YAML/YML/JSON and explicit optional null are supplemented by `docs/adr/ADR-0067-userform-designer-property-metadata.md`.

- JSON geometry editing and the LSP adapter are supplemented by `docs/adr/ADR-0066-userform-designer-geometry-transactions.md`.
- Source-preserving structural add/remove and cascade policy are supplemented by `docs/adr/ADR-0068-userform-designer-structural-control-edits.md`.
- Issue #916 (parent: Issue #914)
- `docs/adr/ADR-0064-userform-designer-text-document-boundary.md`
- `docs/specs/userform-semantic-edit.md`
- `docs/specs/userform-designer.md`
- `docs/specs/ms-oforms.md`
- `internal/vba/userforms/spec/spec.go`
