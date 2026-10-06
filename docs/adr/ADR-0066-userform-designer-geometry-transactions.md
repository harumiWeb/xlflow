# ADR-0066: UserForm Designer Geometry Transactions

## Status

Accepted

## Context

Issues #915 and #916 established a read-only CustomTextEditor and a host-neutral
YAML semantic edit engine. Issue #917 makes geometry interactive and includes
JSON FormSpecs. Pointer events must not flood source or native undo history,
and stale interactions must not overwrite concurrent source edits.

## Decision

- Supplement ADR-0064 with `xlflow/userFormEdit` through the existing selected
  project's LSP. Pass the exact unsaved TextDocument source/version and geometry
  operations; return localized text edits, leaving application to the host.
- Keep drag/resize/key-repeat preview transient. Commit one WorkspaceEdit per
  interaction, including move plus resize for north/west handles. Use native
  CustomTextEditor undo/redo, with no frontend history or automatic save.
- Serialize edits per URI and reject obsolete document/connection results.
  Reject rather than silently rebase an interaction onto a different source.
- Extend ADR-0065 with JSON geometry-only source edits. Preserve original JSON
  token ranges and formatting; share canonical validation and operation rules
  with YAML. Do not broaden JSON structural/property editing in this issue.
- Keep points canonical and centralize screen transforms and container content
  bounds. Restrict UI geometry to parent content areas, protect children when
  shrinking containers, and never repair existing overflow on load. These
  restrictions are UI behavior, not new canonical source validation rules.
- Preserve optional editor registration, normal text editing, read-only
  fallback for preview-only servers and the existing nonce/resource policy.

## Consequences

- Visual edits and source edits share one document and native history. The
  server remains host-neutral and does not write source or call Excel.
- JSON requires a separate lexical writer; unsupported operations and ambiguous
  duplicate keys fail atomically instead of causing whole-document churn.
- Approximate renderer bounds also constrain interaction. Users can author
  layouts outside these bounds in the normal text editor; opening the Designer
  never silently changes those layouts.
- Stale requests must be retried by the user. This favors preserving concurrent
  source edits over interpreting an old pointer gesture against new geometry.

## Alternatives Considered

1. **Rewrite source in the Webview.** Rejected because it duplicates FormSpec
   semantics and destroys source-preserving editing.
2. **Apply each pointer/key-repeat event.** Rejected because it produces noisy
   undo history, excess validation and intermediate authored geometry.
3. **Add an independent CustomDocument/history.** Rejected because the backing
   document is text and its native history already supports source editing.
4. **Serialize all JSON or convert it to YAML.** Rejected because unrelated
   formatting and the user's chosen source format must remain intact.
5. **Automatically reparent or clamp existing source on load.** Rejected because
   opening a Designer must not mutate hierarchy or geometry without an edit.

## Related

- Property transactions and their separate capability are supplemented by `docs/adr/ADR-0067-userform-designer-property-metadata.md`.

- Issue #917 (parent: Issue #914)
- `docs/adr/ADR-0064-userform-designer-text-document-boundary.md`
- `docs/adr/ADR-0065-userform-source-preserving-semantic-edits.md`
- `docs/specs/userform-designer.md`
- `docs/specs/userform-semantic-edit.md`
- `editors/vscode/src/userFormEditor/edits.ts`
- `editors/vscode/test/suite/userFormWebview.test.ts`
- `internal/lspserver/userform_edit.go`
