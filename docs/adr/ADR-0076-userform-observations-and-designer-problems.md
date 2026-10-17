# ADR-0076: UserForm Snapshot Observations and Designer Problems

## Status

Accepted

## Context

UserForm specs combine authored build values with observations captured from
Excel. Synthesizing observations from authored values makes snapshots appear to
contain independent measurements and can turn normalized model data into
source-shaped data. At the same time, existing YAML/YML/JSON captures may carry
genuine observations that source-preserving edits must retain.

Persisted snapshot warnings also need to remain actionable. A generic support
warning for every snapshot-only field adds noise, while the actual warning
entries carry the inspection code and message that should be reviewed in the
Designer.

## Decision

- Keep authored values and observations distinct. Shared normalization does not
  synthesize observed copies from authored values. New file projections and
  Excel snapshots emit authored fields once and do not add observed copies of
  captions, control scalar properties, client dimensions, or TabStrip tab data.
  They retain genuine captured Page geometry and form outer/inside dimensions.
- Continue accepting existing observed data in YAML, YML, and JSON. Localized
  source edits preserve observations already present and never promote them to
  authoring fields or add absent observations.
- Suppress generic `UFV013` support-level warnings only for
  `SupportLevelSnapshotOnly`. Preserve substantive warnings for custom,
  observed-only, and best-effort support. Each persisted warning entry with a
  message becomes a warning issue with its original code and message at
  `warnings[i].message`; the presence of `warnings`, `observed`, or
  `unsupported` containers alone is not a warning.
- Show current issues in a bottom Designer Problems panel that is collapsed by
  default. Its count and structured rows support source navigation and keyboard
  operation, with Escape returning focus to the Designer. Actionable parse,
  validation, and edit errors remain inline at their affected field or
  interaction.

## Consequences

- Snapshots no longer imply that a duplicated authored value was independently
  measured. Real captured Page and form dimensions remain reviewable.
- Older source files keep their observations through local edits without
  changing the meaning of those observations for authoring.
- Persisted warning details remain visible and navigable, while the collapsed
  panel keeps routine Designer work focused on the form.
- Validation must distinguish a generic support-level warning from a
  substantive warning, and source navigation depends on field locations from
  the canonical LSP diagnostics.

## Alternatives Considered

1. **Serialize every normalized observation.** Rejected because it duplicates
   authored values and blurs capture data with authoring intent.
2. **Remove all observed values from existing documents.** Rejected because
   source-preserving edits must retain genuine capture data supplied by users.
3. **Warn whenever a snapshot-oriented container is present.** Rejected because
   the container alone does not describe a problem; persisted warning entries
   and substantive support limitations provide the meaningful issues.
4. **Show warnings only as transient inline text.** Rejected because captured
   warning details need a stable, navigable list while actionable edit errors
   should remain visible at their point of use.

## Related

- [UserForm Designer contract](../specs/userform-designer.md)
- [UserForm semantic edit contract](../specs/userform-semantic-edit.md)
- [UserForm FormSpec reference](../../vitepress/reference/userform-spec.md)
- `internal/vba/userforms/spec/spec.go`
- `internal/vba/userforms/projection/projection.go`
- `internal/vba/userforms/spec/edit/source.go`
