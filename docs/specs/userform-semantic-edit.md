# Source-Preserving UserForm FormSpec Semantic Editing

This specification defines the host-neutral Go API for applying semantic edits
to UserForm FormSpecs while retaining the original source as the canonical
document. Issue #916 provides the YAML API; Issue #917 adds JSON geometry
editing and the Designer LSP adapter; Issue #918 adds scalar property editing
across YAML/YML/JSON and authored-state metadata. The core API remains host-neutral and
does not invoke Excel or provide a CLI command.

## Authority and source model

The API is provided by `internal/vba/userforms/spec/edit`:

```go
func Apply(input spec.SpecInput, source []byte, operations []Operation) (Result, error)
```

`Apply` accepts `.yaml` and `.yml` for all operations, and `.json` for
`moveControl`, `resizeControl`, `setFormProperty` and `setControlProperty`.
Other JSON operations return a structured
unsupported-operation error. `Operation` values are serializable tagged structs.
The supported YAML operation tags are
`setFormProperty`, `setControlProperty`, `moveControl`, `resizeControl`,
`addControl`, `removeControl`, `setParent`, and `reorderControl`.

The canonical `spec.ParseFormSpec` parser and the existing FormSpec validation
functions remain the authority for parsing, normalization, warnings, and
validation. The edit package must not reproduce ID uniqueness, control
property, parent, cycle, Page/MultiPage, or TabStrip validation in an editor or
in a second ruleset.

Source validity and writer capability are separate. `spec.ParseFormSpec`
rejects a source control without an explicit `id` with `UFV004`. Although
`spec.NormalizeFormSpec` can synthesize an ID in an in-memory model, that model
ID does not make the source valid. The edit API does not insert missing IDs or
target normalized-only IDs. Canonical source validation currently does not
reject duplicate control names; the edit API adds no such rule. A source that
parses successfully may still fail stricter compiler or generator checks, and
those downstream checks remain authoritative for writer capability.

Semantic interpretation and source preservation use separate representations:

- The canonical `spec.FormSpec` is the working semantic document used to
  resolve targets and apply operations. Normalization may synthesize observed
  values, control IDs, z-indices, or a flattened control view.
- A YAML syntax representation retains the original bytes, node boundaries,
  scalar spelling, comments, ordering, and whitespace needed to generate local
  source edits. It must not be reconstructed by serializing the normalized
  FormSpec.

This boundary prevents normalization details from becoming authored source.
An edit to one scalar should replace that scalar or insert one mapping entry;
it must not rewrite unrelated controls or the whole document as a recovery
path. Existing comments, whitespace, and order outside the smallest affected
YAML container are preserved. If a localized safe edit cannot be produced,
the operation fails without returning modified source.

Scalar replacement preserves the key, surrounding spacing, and trailing
comment. A missing scalar property is inserted into its owning mapping. Flow
mappings and sequences may be regenerated as the smallest local container when
an edit requires it. If the targeted value depends on an unsafe anchor, alias,
or merge-key expansion, return a structured error; preserve unrelated anchors,
aliases, and merge keys byte-for-byte.

## Operation representation

For JSON geometry and property edits, retain lexical token spans separately from the
canonical model. Replace only targeted scalar values and insert missing
members into the owning object, retaining unrelated bytes, key order,
whitespace and newline style. JSON is never rewritten as YAML or serialized
from the normalized FormSpec. Reject duplicate member keys that would make
source targeting ambiguous, including escaped spellings of the same key.
Canonical parsing, payload validation, atomic failure and numeric no-op
behavior remain shared with YAML. No additional FormSpec validation authority
is introduced by the JSON writer.

Setting a supported `build.*` path may insert the missing `form.build` object;
only the required local object/member is added. Legacy fields and build fields
remain independent. `form.name` changes only that field, without renaming files,
sidecars or VBA references.

The operation is a tagged value with optional payload fields. Pointer fields
distinguish an omitted number from an explicit zero. `Value` is restricted to a
scalar value accepted by the selected FormSpec field. Presence is tracked
independently from a nil value: `ValuePresent` is required for explicit null,
while non-nil values establish presence for existing Go callers. Wire decoding
records whether `value` exists; encoding preserves explicit null.

```go
type OperationType string

type Operation struct {
	Type      OperationType       `json:"type"`
	ControlID string              `json:"controlId,omitempty"`
	Field     string              `json:"field,omitempty"`
	Value     any                 `json:"value,omitempty"`
	ValuePresent bool             `json:"-"`
	Left      *float64            `json:"left,omitempty"`
	Top       *float64            `json:"top,omitempty"`
	Width     *float64            `json:"width,omitempty"`
	Height    *float64            `json:"height,omitempty"`
	Control   *spec.FormSpecControl `json:"control,omitempty"`
	ParentID  string              `json:"parentId,omitempty"`
	Index     *int                `json:"index,omitempty"`
	Cascade   bool                `json:"cascade,omitzero"`
}
```

Fields irrelevant to an operation are ignored only where explicitly stated by
that operation's validation; malformed or contradictory payloads fail with a
structured error. `ControlID` addresses the ID exposed by the current
canonical parse and must match an explicit ID in the source. A missing source
ID fails parsing with `UFV004` before any operation is applied.

| Type                 | Contract                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
| -------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `setFormProperty`    | Set one supported scalar field by an explicit path relative to `form`. Supported paths include `name`, `caption`, `width`, `height`, `build.caption`, `build.width`, `build.height`, `build.clientWidth`, and `build.clientHeight`. The path names the exact field: `caption` never silently redirects to `build.caption`, and build fields never redirect to legacy form fields. `observed.*` is not editable.                                                          |
| `setControlProperty` | Set one supported scalar field named relative to the selected control. Supported fields include `name`, `caption`, `tag`, `controlTipText`, `accelerator`, `text`, scalar `value`, `left`, `top`, `width`, `height`, `tabIndex`, `selectedIndex`, `enabled`, and `visible`. `progId` is not an authorable field. `id` and `type` are immutable. Geometry can be set individually with this operation; `moveControl` and `resizeControl` are the paired-field operations. |
| `moveControl`        | Set both `left` and `top` on `controlId`; both pointers are required.                                                                                                                                                                                                                                                                                                                                                                                                    |
| `resizeControl`      | Set both `width` and `height`; both pointers are required.                                                                                                                                                                                                                                                                                                                                                                                                               |
| `addControl`         | Add exactly one control from `control`. It must have an explicit `id`, `name`, and `type`, and must not contain a nested `controls` tree. `Control.ParentID` supplies the optional parent relationship. `Operation.ParentID` is irrelevant and must be empty; a non-empty value is rejected. The canonical validator checks the resulting control and topology.                                                                                                          |
| `removeControl`      | Remove the target control. `cascade` defaults to false; if descendants remain, an absent or false value rejects removal. `cascade: true` removes the target and its descendants.                                                                                                                                                                                                                                                                                         |
| `setParent`          | Set `parentId`, or make the control a root when it is empty. Optional `left` and `top` replace parent-relative coordinates; omitted values retain the existing local coordinates. Reparenting does not preserve or compute a world-space position.                                                                                                                                                                                                                       |
| `reorderControl`     | Move the target to the zero-based `index` among its current siblings. Sibling order uses stable `zIndex` order, preserving source order for ties. Only sibling `zIndex` values needed for the requested order are updated; `tabIndex` is unchanged.                                                                                                                                                                                                                      |

For property operations a missing `value` is invalid. Explicit null is supported
only for an optional scalar field whose resulting source passes canonical
validation; required fields cannot be null. Null replaces or inserts the
source scalar and never deletes its key. Absent, null, empty string, false and
zero are distinct authored states. An explicit value matching a normalized
default must not be mistaken for an absent-field no-op.

Property removal is unsupported. `observed` fields, property bags,
collections (including lists and tabs), picture data/actions, and non-scalar
property values are not edit targets. Page and TabStrip collection operations
are outside this contract. Existing FormSpec validation still decides whether
a scalar value or resulting control hierarchy is valid.

## Property metadata and authored-state extraction

The host-neutral edit layer exposes supported scalar properties from Go
FormSpec contracts and extracts presence/value from original syntax for the
LSP preview adapter. The Property Grid payload and optional `propertyError`
are specified in [`userform-designer.md`](userform-designer.md). Normalized,
observed and renderer values are not substitutes for authored values. Unsafe
or ambiguous syntax must not produce guessed metadata; the adapter can retain
its valid canonical preview while reporting property metadata unavailable.
No frontend validation rules replace canonical parsing and validation.

## Transaction and identity rules

Operations execute in the supplied order against a private working FormSpec.
The edit is transactional: parsing errors, unsupported operations, target
errors, unsafe source structures, and final-state validation errors return an
error and no result. No intermediate source or model is published. Validate
the complete final state with the canonical FormSpec parser and validators
before producing the result; a sequence may temporarily violate a relationship
that a later operation repairs.

`Apply` parses and validates the source before processing operations. A missing
control `id` therefore fails with `UFV004`, and no operation can repair that
input by inserting or addressing an implicit ID. `NormalizeFormSpec` may
produce model-only IDs for an in-memory value; those IDs are not source targets
for this API.

A semantic no-op returns the original source bytes unchanged and an empty edit
list. For scalar properties this requires matching authored presence and value,
not merely matching a normalized default. The resulting YAML or JSON is parsed
again through `spec.ParseFormSpec`; its normalized
document and validation warnings are returned with the generated source.

## Result and errors

```go
type Result struct {
	Source   []byte
	Document spec.FormSpec
	Edits    []SourceEdit
	Warnings []spec.ValidationIssue
}

type SourceEdit struct {
	Start int    // inclusive UTF-8 byte offset in the original source
	End   int    // exclusive UTF-8 byte offset in the original source
	Text  string // replacement or insertion text
}
```

Every edit range is measured in UTF-8 bytes against the original input, not a
partially edited buffer. Edits are deterministic, ordered by original source
position, and non-overlapping. Applying them to the original source produces
`Result.Source`. Insertions use an empty range (`Start == End`).

Operation failures carry a structured error with the zero-based operation
index, operation type, control ID when relevant, field when relevant, stable
code, message, and suggestion when available. Validation codes such as
`UFVxxx` pass through unchanged from the canonical validator. A rejected
transaction returns no partial `Result`.

## Required implementation tests

The implementation is pure Go and must cover these behaviors:

- Replacing a scalar and inserting a missing property changes only the local
  source span and preserves unrelated comments, spacing, key order, and
  controls. Include CRLF input and multibyte text to verify UTF-8 byte offsets.
- Applying each of the eight operation types, including sequential operations,
  default and disabled cascading removal, parent-relative reparenting, and
  stable sibling reorder without changing `tabIndex`.
- Operation errors must retain their own cause even after an earlier operation
  temporarily invalidates a relationship. Duplicate-ID diagnostics must come
  from the canonical validator, including its message and suggestion; add
  failures retain the ID from the control payload.
- Multiline quoted scalar replacement must leave its trailing comment outside
  the replacement, even when identical comment text occurs inside the value.
  Empty flow sequences replacing the last item of an indentless block sequence
  must be indented beyond the owning key, including nested legacy controls.
- Reorder uses an available integer gap before touching other siblings. When
  no gap exists, assign ordered values in the shortest contiguous sibling range
  containing the moved control, retaining outside values and `tabIndex`. Cover
  ties and integer limits without overflow.
- Flow mappings and sequences regenerate only the smallest required
  container. Targeted unsafe anchors, aliases, and merge keys fail atomically;
  unrelated uses remain unchanged.
- Existing explicit IDs remain stable. Missing source IDs fail with `UFV004`
  before operations; model-only IDs produced by `NormalizeFormSpec` are not
  inserted or addressable. No-ops preserve bytes.
- Duplicate control names are not rejected by canonical source validation and
  must not become an edit-engine validation rule. Verify that writer-specific
  name constraints remain the responsibility of downstream compiler or
  generator checks.
- Final-state validation returns canonical `UFVxxx` issues and never publishes
  a partial source. Error data identifies the failing operation and preserves
  validation codes, messages, and suggestions.
- Explicit `form` paths update only the named legacy or `build.*` field.
  Unsupported observed fields, property removal, property bags, collections,
  pictures, and non-scalar values are rejected.
- YAML/YML inputs support all operations; JSON supports geometry and scalar
  property operations, replacement/insertion including missing `build`, escaped keys, duplicate-key rejection and
  formatting-stable no-ops, with unsupported JSON operations rejected atomically.
- Authored metadata distinguishes absent/null/empty/false/zero without exposing
  observed or renderer defaults as authored values. Cover missing `value`,
  explicit optional null round-trip, required-null rejection and insertion of
  explicit values that happen to match normalized defaults.
- `Result.Edits` are deterministic, non-overlapping original-byte ranges whose
  application equals `Result.Source`; a no-op returns no edits.

These tests require no Excel, COM, or VBE oracle. The API does not change
VBE-facing behavior or establish any Excel runtime claim.

## Out of scope

- JSON structural operations other than geometry and scalar property edits.
- CLI wiring; VS Code/Webview/LSP adapters are specified separately.
- Excel, COM, VBIDE, Designer persistence, or VBE validation.
- Property removal, observed-state authoring, property bags, collection edits,
  and picture editing.
- Recursive control construction in `addControl`, implicit ID insertion or
  targeting, and ID/type changes.

## Related

- Issue #918
- `docs/adr/ADR-0067-userform-designer-property-metadata.md`

- Issue #916
- Issue #917
- `docs/adr/ADR-0066-userform-designer-geometry-transactions.md`
- `docs/adr/ADR-0065-userform-source-preserving-semantic-edits.md`
- `docs/specs/userform-designer.md`
- `docs/specs/ms-oforms.md`
- `internal/vba/userforms/spec/spec.go`
