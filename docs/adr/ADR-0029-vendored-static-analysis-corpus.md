# ADR-0029: Vendored Static-Analysis Corpus Provenance and Synchronization

## Status

Accepted

## Context

Issues #530 and #531 add representative VBA projects to the static-analysis
test corpus. The projects are useful parser and analyzer fixtures, but an
unbounded checkout of an upstream branch would make tests non-reproducible,
would allow source changes to enter the repository without review, and would
make ordinary test runs depend on network availability. The corpus also
contains third-party source and therefore needs an attribution trail that is
separate from the Go dependency inventory.

The corpus is consumed as repository data. It must remain available to normal
tests and future analysis runners even when the upstream repository is offline,
while contributors need a safe, reviewable way to refresh the pinned sources.

## Decision

Vendor the selected projects below the managed
`testdata/static-analysis-corpus/projects/third_party/` tree and describe them
with the schema-versioned
`testdata/static-analysis-corpus/manifest.json`:

- upstream repository `harumiWeb/tree-sitter-vba` at commit
  `c867f27ea3dedc2ccece1eeb0273cdb242899182`;
- `vba-web` from `examples/third_party/VBA-Web-master`, using the `excel`
  profile;
- `vba-json` from `examples/third_party/VBA-JSON-master`, using the
  `generic-vba` profile; and
- `access-examples` from `examples/third_party/Access-examples-master`, using
  the `access` profile.

Each manifest entry records an independent destination path, an explicit
`enabled` value, source origin/path, and provenance (`repository`, SPDX
license, and in-tree `LICENSE` and `SOURCE.md` paths). The manifest is the
single source of truth for the pin and project selection. Its loader is strict:
unknown fields, unsupported schema/profile/origin values, duplicate IDs or
paths, malformed commit hashes, non-canonical or escaping paths, unsorted
project entries, empty notes, and missing attribution metadata are errors.

Refreshes are performed only by the developer-only synchronization script.
The script fetches the exact commit into an isolated temporary checkout,
validates project boundaries and provenance, stages every project, and then
publishes the managed tree atomically. It rejects a local checkout whose `HEAD`
does not equal the manifest commit. Symlinks/reparse points, `.git` content,
special files, missing projects, and dirty managed trees are rejected. A failed
publish restores the previous managed tree and removes only staging/backup
directories owned by the invocation. The script does not rewrite the manifest,
accept branch names, or make network access part of normal tests or production
CLI execution.

Vendoring is preferred over submodules, runtime cloning, or a single merged
workspace. It keeps fixture contents reviewable in the repository, gives each
project a stable boundary/profile, and prevents one project's files or metadata
from silently satisfying another project's fixture contract. A future corpus
runner may consume the manifest and vendored paths, but it must remain a
separate test/developer concern and must not add network or Excel/COM
requirements to production commands.

## Consequences

- Tests are deterministic and offline after the corpus is committed; a pinned
  commit makes a refresh auditable and repeatable.
- A refresh can produce a large source diff and requires preserving each
  project's license and `SOURCE.md`; reviewers must check both code and
  provenance.
- The synchronization script is intentionally more defensive than a generic
  copy operation. Atomic staging and dirty-tree checks prevent partial updates
  but may require a contributor to clean or restore a managed directory before
  retrying.
- Adding or changing a project requires a manifest review, a provenance review,
  and updates to this corpus documentation when the policy changes. The
  existing Go dependency inventory remains unchanged.

## Alternatives Considered

1. **Track an upstream branch or tag** - Rejected because a moving ref makes
   fixtures change without a reviewed manifest diff and can break reproducible
   analysis tests.
2. **Use a Git submodule** - Rejected because submodule initialization is easy
   to omit in test environments and does not provide per-project provenance or
   atomic managed-tree publication.
3. **Clone upstream at test/runtime execution** - Rejected because it makes
   normal tests network-dependent and allows upstream availability or content
   to change a test result.
4. **Merge all examples into one workspace** - Rejected because project
   boundaries, host profiles, and attribution would be lost and duplicate
   names could create analyzer cross-talk.

## Evidence

- Issue #530 (parent) and issue #531 acceptance requirements.
- Corpus contract: `docs/specs/static-analysis-corpus.md`.
- Existing static-analysis ownership and registry boundary:
  `docs/adr/ADR-0013-analyze-runtime-risk-ownership.md` and
  `docs/adr/ADR-0024-shared-static-analysis-rule-registry.md`.
- Developer-only local evidence boundary:
  `docs/adr/ADR-0026-local-vbe-oracle-evidence.md`.
- Pinned manifest and attribution files:
  `testdata/static-analysis-corpus/manifest.json`, each project's `LICENSE`,
  and each project's `SOURCE.md`.

## Supersedes

None.

## Superseded by

None.

## Related

- Issue #530
- Issue #531
- `docs/specs/static-analysis-corpus.md`
- `docs/adr/ADR-0024-shared-static-analysis-rule-registry.md`
- `docs/adr/ADR-0026-local-vbe-oracle-evidence.md`
