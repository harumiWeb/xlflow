# ADR-0030: Isolated Third-Party Corpus Workspaces and Host Profiles

## Status

Accepted

## Context

Issue #530 requires the real-world analyzer corpus to preserve project
boundaries. The vendored projects selected by Issue #531 are ordinary VBA
exports, not xlflow workspaces: their files may be nested, `.cls` can represent
an Excel or Access document module, and the source tree must remain untouched.

The production analyzer already has the correct project-scoped behavior, but
it discovers files from configured xlflow source roots. A corpus-specific
adapter is therefore needed. The corpus also contains Excel, generic VBA, and
Access projects; reporting Excel object-model findings for a non-Excel project
would make the regression evidence misleading.

## Decision

Materialize each enabled third-party manifest project into a unique temporary
xlflow workspace and run the existing production `lint` and `analyze`
implementations against that workspace. Preserve source bytes and relative
paths, keep a complete workspace-to-project source map, and remove only the
invocation-owned temporary child on every exit path.

Use `.bas → standard`, `.cls → class`, and `.frm → form` as conservative
defaults. Add exact, project-relative manifest classifications for document
modules; do not infer document semantics from file names or `Attribute VB_*`
metadata. Unsupported layouts, irregular files, collisions, missing
classifications, and unmapped diagnostics are explicit workspace failures.

Keep host-profile policy inside `internal/staticanalysis/corpus`. The `excel`
profile uses the normal analyzer configuration. `generic-vba` and `access`
remove Excel object-model diagnostics from corpus evidence, disabling
configurable rules in generated configs and filtering the two always-on
Excel-member rules at normalization. The shared rule registry and production
Analyzer API remain unchanged.

Normalized third-party diagnostics use `third_party/<manifest-id>` as the
project identity and the original project-relative file path. Disabled
projects require a manifest note and are reported as documented skips.

## Consequences

- Existing parser, symbol, CFG, and call-resolution semantics are reused
  without rewriting vendored source or creating a merged synthetic project.
- Temporary workspace creation and source mapping add a cleanup and failure
  boundary that must be tested separately from analyzer diagnostics.
- Exact classification entries require manifest maintenance when an upstream
  project adds document modules, but they prevent unsafe host-specific guesses.
- Profile filtering keeps non-Excel evidence useful while intentionally making
  the policy corpus-specific; future host taxonomy work can extend this layer
  without changing production rule metadata.

## Alternatives Considered

1. **Merge all third-party files into one workspace** - Rejected because
   project-local symbols and host semantics would cross-contaminate results.
2. **Modify vendored sources into committed xlflow layouts** - Rejected because
   it breaks upstream byte preservation and makes refreshes harder to audit.
3. **Infer document modules from names or VB attributes** - Rejected because
   exported `.cls` conventions are not reliable across VBA hosts.
4. **Add host metadata to the shared rule registry** - Rejected for v1 because
   host selection is a corpus evidence concern and would expand production
   API/registry compatibility scope.

## Evidence

- Issue #530 parent acceptance requirements and Issue #533 adapter scope.
- `docs/specs/static-analysis-corpus.md`.
- `internal/staticanalysis/corpus/manifest.go` and `runner.go`.
- `testdata/static-analysis-corpus/manifest.json` classifications.
- `docs/adr/ADR-0024-shared-static-analysis-rule-registry.md`.

## Supersedes

None.

## Superseded by

None.

## Related

- Issue #530
- Issue #533
- Issue #534
- `docs/adr/ADR-0029-vendored-static-analysis-corpus.md`
- `docs/specs/static-analysis-corpus.md`
