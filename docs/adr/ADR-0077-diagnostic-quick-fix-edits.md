# ADR-0077: Diagnostic Quick Fix Edits

## Status

Accepted

## Context

The existing VB044 correction and documentation actions carry one replacement.
Adding more diagnostic corrections requires one representation for multiple
edits and diagnostic association. Reimplementing corrections in VS Code and a
future CLI would duplicate applicability and source-coordinate rules.

## Decision

- Keep action generation in protocol-neutral Go code. `intel.CodeAction`
  contains a list of document-local edits and optional associated diagnostics;
  LSP maps these to `WorkspaceEdit` and `CodeAction.diagnostics`.
- All edits in an action refer to the same captured source revision. Existing
  versioned edits, cancellation, and lifecycle validation remain authoritative.
  Legacy clients retain their existing unversioned-edit fallback.
- Keep rule detection and correction applicability with their Go diagnostic owner:
  syntax corrections in lint, TypeDB-backed assignment corrections in intel.
  Reuse the assignment diagnostic predicate on selected candidate lines rather
  than moving binding logic into lint or duplicating it in an editor adapter.
  Actions evaluate only enabled correction providers, apply normal inline
  suppression, and avoid whole-file lint and ProcedureIR/CFG construction.
- Use registry `fix_available` for implemented correction capability. It does
  not authorize Fix All or imply that every occurrence can be corrected.
- Start with VB044 and opt-in VB084. Only remove a literal `Step 1` clause
  whose surrounding CST tokens and horizontal whitespace prove a local edit.
  Preserve comments and unrelated continuation markers; withhold a correction
  when the removed span would cross a physical line continuation.
- Add document Fix All through an explicit whitelist of VB044, VB080, VB084,
  and VB086, reusing individual syntax correction providers. Keep VB001 and
  type-backed Set corrections individual-only because their intent and
  preconditions need separate consideration. `fix_available` is insufficient
  authorization for bulk edits.
- Require explicit source-kind requests for aggregation so routine lightbulb
  requests do not scan the whole document. Parent source-kind requests yield
  separate actions per rule. Automatic requests omit aggregation.
- Require an open, versioned document and `documentChanges` capability for Fix
  All. Reject an entire rule action on duplicate, overlapping or invalid edits
  rather than silently selecting an order-dependent subset. This preserves
  one captured revision and an atomic editor transaction.

## Consequences

- Adapters can expose multiple edits without learning VBA transformations.
- Additional corrections can reuse diagnostic association and suppression.
- Conservative applicability leaves some diagnosed occurrences without a fix.
- Legacy clients retain individual corrections but cannot request unversioned
  bulk edits. Explicit eligibility limits bulk coverage in exchange for
  conservative applicability and bounded discovery cost.
- Cross-file preconditions and CLI application still require separate contracts.

## Alternatives Considered

1. **Add individual VS Code providers.** Rejected because applicability would
   diverge from the CLI/language-server analysis and duplicate VBA knowledge.
2. **Retain one replacement and replace a larger source span.** Rejected
   because future corrections may need independent edits that preserve trivia.
3. **Run complete diagnostics for every action request.** Rejected because
   automatic editor requests would repeat unrelated analysis on large files.

## Related

- [Quick Fix contract](../specs/diagnostic-quick-fixes.md)
- [Reusable LSP](ADR-0014-reusable-vba-lsp-server.md)
- [Rule registry](ADR-0024-shared-static-analysis-rule-registry.md)
- [Roadmap #944](https://github.com/harumiWeb/xlflow/issues/944)
- [Initial implementation #945](https://github.com/harumiWeb/xlflow/issues/945)
- [Document Fix All #952](https://github.com/harumiWeb/xlflow/issues/952)
