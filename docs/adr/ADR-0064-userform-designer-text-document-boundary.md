# ADR-0064: UserForm Designer TextDocument Boundary

## Status

Accepted

## Context

Issue #915 adds a visual editor for canonical UserForm FormSpecs. Users may
have unsaved YAML or JSON edits while the visual preview is open, so the
preview must follow the document in VS Code without becoming a second place
where the form is authored. Parsing and validation must also match xlflow's
canonical FormSpec contract.

The extension needs a rendering model for nested controls and transient error
states. The Webview must receive source data safely and stay within a narrow
capability boundary while the visual editor remains read-only.

## Decision

- Keep the VS Code `TextDocument` as the sole canonical source. Register
  `xlflow.userFormDesigner` as an optional custom text editor for YAML, YML,
  and JSON candidates, then accept only FormSpecs directly under the configured
  forms root's `specs` directory.
- Send the current document URI, version, and full text to the selected
  project's xlflow LSP through the experimental `xlflow/userFormPreview`
  request. The LSP parses and validates that text with the canonical FormSpec
  parser and returns the same version with an optional document, warnings, or
  error. Discard replies for an obsolete document version or LSP connection.
- Use the LSP instead of starting a CLI process. The LSP already has the
  selected project's configuration and canonical parser. A CLI preview path
  would need another invocation contract and a way to transport unsaved buffer
  text, creating extra process and synchronization paths.
- Render the response in a read-only Preact and TypeScript Webview. The host
  owns document access and sends typed messages; the Webview can request the
  text editor to open but cannot send edits. Keep the last valid rendering
  visible while reporting a current parse, validation, project, or connection
  error.
- Use Preact rather than a hand-built DOM renderer. The component model keeps
  recursive container and control rendering declarative while keeping the
  preview implementation separate from VS Code APIs. Vanilla TypeScript would
  avoid a UI dependency, but would make nested rendering and updates more
  imperative without improving the source or security boundary.
- Keep rendering intentionally approximate. Use display defaults for absent
  geometry, derive Page layout from its parent MultiPage, and do not implement
  editing, zoom, Page selection, or image asset loading in this stage.
- Apply a strict nonce Content Security Policy and restrict local Webview
  resources to the packaged Designer bundle.

## Consequences

- Unsaved edits in the normal text editor are the preview input, and the
  canonical parser remains the single authority for FormSpec interpretation.
- Users can keep the normal YAML or JSON editor and explicitly reopen an
  eligible document with the visual Designer.
- An LSP that does not advertise `userFormPreview`, a missing LSP connection,
  or a document outside the selected project cannot provide the preview.
- The preview is useful for layout inspection but does not promise exact VBA
  Designer fidelity. Missing geometry and images are represented with display
  defaults or placeholders.
- The request and Webview protocols add versioned preview-specific models.
  Future visual editing requires an explicit edit protocol that updates the
  canonical `TextDocument`; it must not introduce a persisted parallel model.
- Preact adds a Webview dependency and bundle, while the nonce policy and
  resource restrictions limit the Webview's execution and loading surface.

## Alternatives Considered

1. **Run a CLI command to parse the preview.** Rejected because a command would
   add a process invocation and require a separate way to pass unsaved source.
   The selected project's LSP already owns the canonical parser and receives
   the exact current buffer text.
2. **Parse FormSpec in the extension host.** Rejected because it would
   duplicate canonical parsing and validation behavior in TypeScript.
3. **Build the Webview with vanilla TypeScript and DOM APIs.** Rejected because
   nested controls, parent containers, and refreshable error/document states
   are clearer as components; this stage accepts the small dependency and
   bundle cost of Preact.
4. **Make the Designer the default editor or add visual editing now.** Rejected
   because normal text editing must remain directly available and this stage
   establishes only a read-only preview boundary.

## Related

- Interactive geometry is supplemented by `docs/adr/ADR-0066-userform-designer-geometry-transactions.md`.
- Issue #915
- `docs/specs/userform-designer.md`
- `docs/specs/ms-oforms.md`
- `editors/vscode/src/userFormEditor/provider.ts`
- `editors/vscode/src/userFormEditor/protocol.ts`
- `internal/lspserver/userform_preview.go`
