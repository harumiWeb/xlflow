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

## Proposed CLI Extension (Draft)

### Context

The accepted editor contract already provides rule-scoped corrections for one
document revision. A source-only CLI can reuse that behavior for explicit
single-file application, but it needs a filesystem boundary and a publication
contract that editor buffers do not have.

### Proposed Decision

- Add `xlflow fix <path> --rule VB044|VB080|VB084|VB086 [--apply] [--json]`.
  It requires the current project's `xlflow.toml`, accepts one file inside
  configured modules, classes, or workbook source roots or the test root, and
  initially supports only `.bas` and `.cls`. Under the configured forms root,
  permit only `forms/code/*.bas` when `[userform].code_source = "sidecar"`;
  reject forms assets, specs, and inactive sidecar paths regardless of file
  extension. Exclude `.frm` until its `.frx` sidecar has a transaction
  contract.
- Reuse the shared `intel` same-rule correction provider, the existing rule
  configuration and suppression behavior, and UTF-16 range application. Do
  not load TypeDB, run whole-project lint, or construct ProcedureIR/CFG.
- Make preview the default. `--apply` recaptures the current file and computes
  its edits again; preview results are not persisted or reused. Require UTF-8
  without BOM and expose project-relative slash-separated paths, uppercase
  rule IDs, changed/applied state, source/result SHA-256 values, edits,
  associated diagnostics, and a unified diff in JSON.
- Publish a changed file through a same-directory staging file and native
  atomic replace. Reject symlink targets and non-regular files; reject a
  read-only target only when a changed result will be applied. Recheck target
  identity, mode, and contents before publication and clean up staging files
  on pre-publication failure or cancellation.
- Treat conflict detection as best effort. It does not promise a race-free
  compare-and-swap against arbitrary external writers or crash durability.
  Keep workbook bridges, COM, and unsaved editor buffers outside the command.
- Advertise source-tree mutation to consumers through policy
  `ResourceSourceTree`, operation `mutate`, and `ParallelSafe: false`. This
  capability metadata acquires no source-tree lock or workbook lease and
  provides no `--wait` behavior; external-writer races remain best-effort.

### Consequences

- CLI callers and editors use one correction implementation and coordinate
  contract, while explicit preview remains the default for filesystem edits.
- The initial command supports fewer file kinds than project-wide source
  discovery. Extending it to UserForms or multiple files requires a separate
  transaction decision.
- Hashes describe each invocation's captured and proposed bytes; they are not
  durable plan identifiers and do not bind a later apply to an earlier
  preview.

## Proposed Individual VB020 Extension (Draft)

### Context

VB020's procedure-local reference predicate excludes assignment-target
occurrences. Its warning therefore does not establish that deleting a local
declaration alone is safe: an assigned-but-unread variable can still have
writes. The shared correction model already supports individual diagnostic-
associated edits, while bulk eligibility remains an explicit separate policy.

### Proposed Decision

- Add one individual action, **Remove unused local variable**, only when the
  existing enabled VB020 predicate reports the diagnostic and a separate
  occurrence check finds no other same-name identifier in the procedure,
  including writes. Reuse the published diagnostic and its normal configuration
  and suppression behavior; a diagnostic alone does not authorize deletion.
  Bound the additional check to the current parsed procedure; do not rerun
  whole-file lint or construct ProcedureIR/CFG for this action.
- Restrict removal to a standalone, single-declarator procedure-local `Dim`
  with an explicit primitive scalar type: `Boolean`, `Byte`, `Integer`, `Long`,
  `LongLong`, `LongPtr`, `Single`, `Double`, `Currency`, `Date`, `String`, or
  `Variant`. Exclude fixed-length strings, constants, `Static` declarations or procedures, `As New`,
  objects, arrays, UDTs, multi-declarators, colon-separated or continued
  statements, conditional-compilation branches, recovered syntax, and
  ambiguous binding. Withhold ambiguous same-name occurrences conservatively.
- Remove only the declaration statement's CST byte span. Keep indentation,
  trailing comments, line endings, blank lines, and all unrelated source bytes
  unchanged.
- Keep VB020 default-enabled and keep its current diagnostic predicate,
  configuration, and suppression policy. Registry `fix_available` advertises
  provider availability only; it does not enable VB020 or authorize Fix All.
  Keep document Fix All and CLI eligibility limited to VB044, VB080, VB084,
  and VB086, including no aggregate action for an explicit
  `source.fixAll.xlflow.vb020` request.
- Keep `quickfix-unused-local-before` and `quickfix-unused-local-after` as
  unbound maintainability observations. Both compile in VBE after the known
  accept/reject controls with confirmed cleanup. Compilation supports syntax
  acceptance; deletion safety additionally depends on the restricted types
  and absence of other occurrences.

### Consequences

- Assigned-but-unread locals and other same-name occurrences cannot be removed
  based solely on VB020's unread-reference diagnostic.
- The syntax and binding gates intentionally leave many diagnosed declarations
  without an action; `fix_available` remains a capability flag, not a promise
  that every VB020 occurrence can be corrected.
- The proposal extends individual correction coverage without changing the
  accepted edit ownership, revision, suppression, or explicit bulk-eligibility
  decisions above.

## Alternatives Considered

1. **Add individual VS Code providers.** Rejected because applicability would
   diverge from the CLI/language-server analysis and duplicate VBA knowledge.
2. **Retain one replacement and replace a larger source span.** Rejected
   because future corrections may need independent edits that preserve trivia.
3. **Run complete diagnostics for every action request.** Rejected because
   automatic editor requests would repeat unrelated analysis on large files.

## Related

- [Quick Fix contract](../specs/diagnostic-quick-fixes.md)
- [CLI contract](../specs/cli-contract.md)
- [Reusable LSP](ADR-0014-reusable-vba-lsp-server.md)
- [Rule registry](ADR-0024-shared-static-analysis-rule-registry.md)
- [Roadmap #944](https://github.com/harumiWeb/xlflow/issues/944)
- [Initial implementation #945](https://github.com/harumiWeb/xlflow/issues/945)
- [Document Fix All #952](https://github.com/harumiWeb/xlflow/issues/952)
- [CLI application #954](https://github.com/harumiWeb/xlflow/issues/954)
- [Individual VB020 correction #956](https://github.com/harumiWeb/xlflow/issues/956)
