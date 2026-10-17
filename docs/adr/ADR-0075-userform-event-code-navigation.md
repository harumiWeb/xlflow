# ADR-0075: UserForm Event Code Navigation

## Status

Accepted

## Context

ADR-0067 makes the running LSP authoritative for Designer metadata, and
ADR-0073 keeps layout edits in a deferred source draft. Event navigation crosses
that boundary: a control's persisted name determines a procedure in a separate
VBA document. Reusing layout history or generating from an unsaved name could
leave code detached from the saved form.

The runtime TypeDB does not retain all information needed for VBE procedure
generation, including parameter passing modes and visible default events.
TypeLib presence alone is insufficient: Image exposes hidden Click/DblClick
members that Excel's VBE rejects for CreateEventProc.

## Decision

- Keep event applicability, default events and exact VBA declarations in a
  dedicated Go catalogue verified against local VBE. Serve it through a
  separately negotiated `userFormEvents: 1` capability and explicit list/edit
  requests. Neither Webview nor extension maintains a second signature table.
- Keep single-click selection. Double-click opens the default event; an explicit
  event action offers the catalogue through a native picker. Existing handlers
  are navigation targets, not templates to rewrite.
- Require explicit publication and saving of the targeted FormSpec before code
  generation when its source or Designer draft is dirty. Revalidate target
  identity after saving and after asynchronous selections.
- Respect `[userform].code_source`: edit the sidecar in sidecar mode and the
  existing exported form's code section in frm mode. Do not migrate modes,
  synchronize another artifact, or rename files as a side effect.
- Let the host own file resolution, open buffers, revision checks and native
  WorkspaceEdit application. Let Go own parsing, collision detection, minimal
  insertion planning and UTF-16 code positions. Apply code changes as one native
  code-document undo step, separate from Designer layout history.
- Generate code locally without Excel, automatic code saving or a workbook
  build. Excel/VBE is development evidence, not a runtime dependency.

## Consequences

- Old servers keep existing Designer functions; event actions need the new
  capability. Catalogue changes require focused VBE evidence rather than
  frontend guesses or machine-dependent TypeLib generation at runtime.
- Dirty code buffers remain authoritative, and repeated requests must be
  serialized per code document. Ambiguous declarations, stale revisions and
  inconsistent form identities stop generation instead of overwriting code.
- A completed FormSpec save remains saved if the user later cancels the event
  picker. The resulting code buffer remains unsaved until its ordinary save.
- Form build continues to consume sidecar code regardless of frm display mode.
  Event editing does not change that command contract; frm users must use their
  existing frm workflow or explicitly migrate before building canonical forms.
- Control rename, deletion and appearance/hierarchy copying do not rewrite,
  remove or transplant event handlers. Custom ActiveX signatures remain outside
  the initial catalogue.

## Alternatives Considered

1. **Generate in the Webview.** Rejected because signature and applicability
   tables would drift from the installed LSP and its VBA parser.
2. **Use generated TypeDB members directly.** Rejected because default-event,
   hidden-member and parameter-passing information is incomplete for this use.
3. **Generate through live VBIDE.** Rejected because source editing should work
   without opening Excel or mutating a workbook.
4. **Always create a sidecar.** Rejected because it would silently change the
   authoritative code source of existing frm projects.

## Related

- [Designer contract](../specs/userform-designer.md)
- [Deferred save lifecycle](ADR-0073-userform-designer-deferred-save-document-lifecycle.md)
- [Property metadata authority](ADR-0067-userform-designer-property-metadata.md)
- `internal/vba/userforms/events/`
- `internal/lspserver/userform_events.go`
- `editors/vscode/src/userFormEditor/events.ts`
