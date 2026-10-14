# ADR-0073: UserForm Designer Deferred-Save Document Lifecycle

## Status

Accepted

## Context

ADR-0064 established an optional `CustomTextEditor` whose preview follows the
open `TextDocument`. ADR-0066 extended that boundary so Designer operations
apply `WorkspaceEdit` changes directly to the same document and use its native
undo history without automatic save. Later Designer capabilities retained
that live-edit lifecycle.

The approved deferred-save workflow needs a source-text draft shared by every
Designer panel for one FormSpec URI. Designer operations must remain
source-preserving and use xlflow LSP parsing and validation, while leaving the
open source document unchanged until a VS Code save action. Edits made in the
regular text editor must still appear immediately in the Designer. When those
edits overlap a dirty Designer draft, silently choosing either version would
lose user work.

VS Code must continue to own undo, save, AutoSave, Save As, revert, and backup
behavior. The Webview must not become a second FormSpec parser, validator, or
canonical editing model.

## Decision

- Register the optional Designer through VS Code `CustomEditorProvider`, while
  keeping the normal text editor available for each eligible YAML, YML, and
  JSON FormSpec.
- Keep one `CustomDocument` and one source-text draft per FormSpec URI. All
  Designer panels for that URI use the same draft and history. The draft starts
  from the latest open `TextDocument` text and preserves the source text,
  including comments and formatting. It is the pending edit buffer; the
  `TextDocument` remains the project source that receives saved changes.
- Give the Designer a dedicated `xlflow-userform-designer` editor URI mapped
  to the original file URI. VS Code scopes native Undo by resource; sharing the
  file URI lets text-editor edits and save-time `WorkspaceEdit` entries enter
  the Designer's native Undo stack. Opening a file with the Designer redirects
  to its dedicated editor resource. LSP requests, source synchronization and
  edit queues retain the original file URI.
- Preview and validate the exact current draft through the existing xlflow LSP
  requests. The LSP and canonical FormSpec parser remain the authority for
  semantic interpretation and validation. The host applies localized LSP text
  edits to the shared draft; the Webview does not implement independent
  FormSpec rules.
- Represent each completed Designer action as a `CustomDocumentEditEvent`.
  VS Code native Undo/Redo changes the shared draft, and every Designer panel
  for the URI observes the same history. The draft engine owns rebasing this
  history when regular text-editor changes are merged into the draft.
- Keep Designer edits out of the open `TextDocument` until a save action.
  Regular text-editor changes are observed immediately and reflected in the
  preview. If the Designer draft is clean, adopt the current text-editor
  contents. If it is dirty, merge non-overlapping changes against the last
  synchronized source and rebase draft history.
- When a dirty draft meets source changes to control order, identity, type or
  parent, require explicit conflict resolution. Compare canonical LSP
  projections before merging; a textual merge alone can attach a geometry
  edit to another control after a source reorder.
- If overlapping changes cannot be merged safely, retain both the current
  text-editor source and the Designer draft, show a source conflict, and
  disable saving until the user resolves it. Provide explicit localized
  `Use YAML/JSON changes` and `Use Designer draft` choices. When a granular
  merge is unavailable, selecting a whole-document version requires explicit
  user confirmation. Never discard either side through an automatic
  last-writer-wins decision.
- Save only after reconciling against the latest open `TextDocument`. Apply
  the resolved source edits through VS Code document APIs to that current
  document, then call `TextDocument.save()`. Do not write the workspace file
  directly or overwrite a dirty text-editor buffer from an older snapshot.
  Save As follows the same VS Code document-save path for the destination.
- Use the normal VS Code custom-document lifecycle for backup, restore, and
  revert. Revert discards unsaved Designer edits through the custom-document
  lifecycle. VS Code AutoSave invokes the same save path according to the
  user's standard AutoSave configuration; the extension does not override
  that setting.

## Consequences

- FormSpec source (YAML, YML, and JSON) remains unchanged by Designer actions
  until Save or configured AutoSave. The regular text editor can still change
  the shared draft input immediately, and non-overlapping edits survive both
  editing paths.
- Native VS Code history now belongs to the shared custom document. External
  source changes require the draft engine to rebase that history correctly;
  this adds state and merge work to the provider lifecycle.
- Conflicts can stop a save until the user chooses or composes a valid source.
  Whole-document resolution is explicit and may discard one conflicting
  version only after the user chooses it.
- Save and AutoSave must reconcile and apply edits to the newest open
  `TextDocument` before saving it. A failed apply or save leaves the draft
  available for correction or retry.
- Backup, restore, revert, and Save As follow VS Code custom-editor lifecycle
  behavior, which adds provider responsibilities compared with the former
  write-through text-editor integration.
- Existing LSP source-edit, canonical validation, stale-response, and
  source-preservation contracts continue to apply to the draft.

## Alternatives Considered

1. **Continue applying every Designer edit to the `TextDocument`.** Rejected
   because it writes through before an explicit save and cannot provide the
   approved deferred draft lifecycle.
2. **Keep draft state and undo history in each Webview.** Rejected because
   panels could diverge and history would bypass VS Code native Undo/Redo.
3. **Save the draft directly to the workspace file.** Rejected because this
   can overwrite a newer dirty text-editor buffer and bypasses VS Code save,
   AutoSave, backup, and Save As behavior.
4. **Choose one side automatically when external and Designer edits overlap.**
   Rejected because either last-writer-wins direction can lose user-authored
   source.

## Related

- `docs/adr/ADR-0064-userform-designer-text-document-boundary.md`
- `docs/adr/ADR-0065-userform-source-preserving-semantic-edits.md`
- `docs/adr/ADR-0066-userform-designer-geometry-transactions.md`
- `docs/specs/userform-designer.md`
- `editors/vscode/src/userFormEditor/provider.ts`
- `editors/vscode/test/suite/userFormProvider.test.ts`
- `editors/vscode/test/suite/userFormEditor.test.ts`

## Supersedes

- The direct live-`TextDocument` edit target and `CustomTextEditor` lifecycle
  decisions in ADR-0064.
- The `WorkspaceEdit`-backed TextDocument history and no-automatic-save
  assumptions in ADR-0066.

## Superseded by

- None
