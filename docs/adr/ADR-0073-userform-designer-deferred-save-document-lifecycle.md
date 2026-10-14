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
source-preserving and use xlflow LSP parsing and validation. Native VS Code
save callbacks, including AutoSave, Save All, and menu Save, must durably
checkpoint the draft without publishing it to the open source document. Only
the Designer's explicit Save action publishes the validated draft to YAML,
YML, or JSON. Edits made in the regular text editor must still appear
immediately in the Designer. When those edits overlap a pending Designer draft,
silently choosing either version would lose user work.

VS Code must continue to own native Undo/Redo, custom-document save callbacks,
AutoSave, Save As, Revert, backup, and hot-exit behavior. The Webview must not
become a second FormSpec parser, validator, or canonical editing model.

## Decision

- Register the optional Designer through VS Code `CustomEditorProvider`, while
  keeping the normal text editor available for each eligible YAML, YML, and
  JSON FormSpec.
- Keep one `CustomDocument` and one source-text draft per FormSpec URI. All
  Designer panels for that URI use the same draft and history. The draft starts
  from the latest open `TextDocument` text and preserves the source text,
  including comments and formatting. Persist pending drafts in extension
  internal storage keyed by the original FormSpec URI so they survive normal
  close, reopen, and extension restart. The `TextDocument` remains the project
  source; checkpointing never writes it.
- Give the Designer a dedicated `xlflow-userform-designer` editor URI mapped
  to the original file URI. VS Code scopes native Undo by resource, so Designer
  actions retain native custom-document Undo/Redo while regular text-editor
  edits keep their source-document history. Opening a file with the Designer
  redirects to its dedicated editor resource. LSP requests, source
  synchronization, and publication retain the original file URI.
- Preview and validate the exact current draft through the existing xlflow LSP
  requests. The LSP and canonical FormSpec parser remain the authority for
  semantic interpretation and validation. The host applies localized LSP text
  edits to the shared draft; the Webview does not implement independent
  FormSpec rules.
- Represent each completed Designer action as a `CustomDocumentEditEvent`.
  VS Code native Undo/Redo changes the shared draft, and every Designer panel
  for the URI observes the same history. Undo/Redo never publishes source.
  The draft engine owns rebasing this history when regular text-editor changes
  are merged into the draft.
- Keep Designer edits out of the open `TextDocument` until explicit Designer
  publication. Native save callbacks only checkpoint internal storage.
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
  disable explicit source publication until the user resolves it. Native
  checkpoint saves remain available. Provide explicit localized
  `Use YAML/JSON changes` and `Use Designer draft` choices. When a granular
  merge is unavailable, selecting a whole-document version requires explicit
  user confirmation. Never discard either side through an automatic
  last-writer-wins decision.
- The native `CustomEditorProvider` save callback checkpoints the current
  draft durably to extension internal storage only. This applies to AutoSave,
  File Save All, and menu Save. It never writes YAML, YML, or JSON. The native
  tab dirty marker indicates whether the internal checkpoint is current; a
  separate `YAML/JSON not applied` badge indicates that the draft has not been
  published to source.
- Publish only through the explicit Designer Save button or the
  Designer-scoped Ctrl+S/Cmd+S command. Validate the current draft through the
  existing LSP, reconcile it against the newest open `TextDocument`, apply the
  resolved source edits through VS Code document APIs, then call
  `TextDocument.save()`. Do not write the workspace file directly or overwrite
  a newer dirty text-editor buffer. After a successful publish, refresh the
  internal checkpoint. Save As is an explicit publish to the chosen
  destination through its current `TextDocument` and document-save path.
- Use the normal VS Code custom-document lifecycle for backup, restore,
  hot-exit, and Revert. Revert discards the pending Designer draft and adopts
  the latest source `TextDocument` buffer, including unsaved text-editor
  changes. AutoSave follows the user's standard VS Code configuration and
  checkpoints only; the extension does not override global AutoSave settings.
- A corrupt checkpoint produces a visible recovery error and remains stored;
  it is never silently discarded or overwritten. Clear pending recovery state
  only after an explicit source publish or Revert completes successfully.

## Consequences

- Native Save, Save All, and AutoSave durably checkpoint the Designer draft
  without changing FormSpec source. The native tab dot can therefore clear
  while the independent `YAML/JSON not applied` badge remains. Only explicit
  Designer Save publishes source. Regular text-editor changes still appear
  immediately in the preview, and non-overlapping edits survive both paths.
- Native VS Code history now belongs to the shared custom document. External
  source changes require the draft engine to rebase that history correctly;
  this adds state and merge work to the provider lifecycle.
- Conflicts disable explicit Designer Save until the user chooses or composes a
  valid source; native checkpoint saves remain available. Whole-document
  resolution is explicit and may discard one conflicting version only after
  the user chooses it.
- Explicit Designer Save must validate and reconcile against the newest open
  `TextDocument` before publishing. A failed apply or save leaves the pending
  draft and recovery state available for correction or retry. Save As explicitly
  publishes to its chosen destination.
- URI-keyed extension checkpoints preserve drafts across close, reopen, and
  restart. Backup, restore, hot-exit, and Revert remain supported. Corrupt
  checkpoints require visible recovery and remain available until a successful
  explicit publish or Revert resolves them.
- Existing LSP source-edit, canonical validation, stale-response, and
  source-preservation contracts continue to apply to the draft.

## Alternatives Considered

1. **Continue applying every Designer edit to the `TextDocument`.** Rejected
   because it writes through before an explicit save and cannot provide the
   approved deferred draft lifecycle.
2. **Keep draft state and undo history in each Webview.** Rejected because
   panels could diverge and history would bypass VS Code native Undo/Redo.
3. **Publish source from the native custom-document save callback.** Rejected
   because AutoSave, Save All, or menu Save would apply Designer edits to YAML,
   YML, or JSON without the explicit Designer publish action.
4. **Write the workspace file directly from the Designer.** Rejected because
   this can overwrite a newer dirty text-editor buffer and bypasses VS Code
   document-save, backup, and Save As behavior.
5. **Choose one side automatically when external and Designer edits overlap.**
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
