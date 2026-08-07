# Static-analysis corpus

This specification defines the checked-in VBA projects used as static-analysis
fixtures and the developer-only operation that refreshes them. The corpus is
test data, not a runtime dependency of `xlflow`.

## Location and pinned source

The corpus root is `testdata/static-analysis-corpus`. Its manifest is
`manifest.json`; vendored projects live below
`projects/third_party/<project-id>`. The initial manifest is schema version 1
and pins:

| ID                | Profile       | Upstream project path                         | Destination                            |
| ----------------- | ------------- | --------------------------------------------- | -------------------------------------- |
| `access-examples` | `access`      | `examples/third_party/Access-examples-master` | `projects/third_party/access-examples` |
| `vba-json`        | `generic-vba` | `examples/third_party/VBA-JSON-master`        | `projects/third_party/vba-json`        |
| `vba-web`         | `excel`       | `examples/third_party/VBA-Web-master`         | `projects/third_party/vba-web`         |

All three entries currently have `enabled: true`. Synchronization reproduces
every manifest entry, including an explicitly disabled entry; a future test
runner may use `enabled` to select cases without changing the vendored tree.

The upstream is `harumiWeb/tree-sitter-vba` at the full 40-character commit
`c867f27ea3dedc2ccece1eeb0273cdb242899182`. A branch, tag, abbreviated hash,
or floating dependency is not a valid pin.

## Manifest schema v1

The JSON document has exactly these top-level fields:

```json
{
  "schema_version": 1,
  "upstream": {
    "repository": "harumiWeb/tree-sitter-vba",
    "commit": "c867f27ea3dedc2ccece1eeb0273cdb242899182"
  },
  "projects": [
    {
      "id": "vba-web",
      "path": "projects/third_party/vba-web",
      "profile": "excel",
      "enabled": true,
      "notes": "optional non-empty reviewer note",
      "source": {
        "origin": "tree-sitter-vba",
        "path": "examples/third_party/VBA-Web-master"
      },
      "provenance": {
        "repository": "https://github.com/VBA-tools/VBA-Web",
        "license": "MIT",
        "license_file": "LICENSE",
        "source_file": "SOURCE.md"
      }
    }
  ]
}
```

`notes` is optional; when present it must contain non-whitespace text without
leading or trailing padding. The allowed profiles are `excel`, `generic-vba`,
and `access`. `source.origin` is
currently only `tree-sitter-vba`. `provenance.license` is an SPDX identifier.
The `license_file` and `source_file` values are required relative paths inside
the destination project and must be present in the checked-in tree. `SOURCE.md`
records the original project, repository, license, fixture purpose, and any
normalization or file-removal performed during import.

The loader uses strict JSON decoding. It rejects unknown fields, unsupported
schema/profile/origin values, missing required values, duplicate IDs or
destination/source paths, duplicate provenance paths, malformed repository
names or commit hashes, overlapping project paths, and unsorted project
entries. IDs are stable and lowercase; projects are ordered lexicographically by
ID. All paths use `/`,
are relative, contain no empty, `.` or `..` segment, do not name a drive/UNC
root, and remain within the corpus or upstream checkout boundary. A project
destination must remain under `projects/third_party`.

## Provenance and attribution

Every project keeps its upstream `LICENSE` and a repository-local `SOURCE.md`.
The source file is metadata, not an analyzer fixture, and must not be omitted
when the project is refreshed. The imported Access examples intentionally keep
only exported VBA source files and document removed binaries; VBA-Web and
VBA-JSON document that no source modifications were made apart from permitted
line-ending normalization. `THIRD_PARTY_LICENCES.md` lists these three MIT
projects separately from the Go module dependency inventory.

Changing a pin or adding a project requires reviewing the upstream repository,
license, source metadata, source path, profile, and resulting tree together.
Do not silently replace an upstream project's license with the repository's
license or rely on a URL without preserving the in-tree metadata.

## Synchronization contract

`scripts/dev/sync-static-analysis-corpus.ps1` is the only supported refresh
entry point. It reads the manifest as its sole source of project selection and
the commit pin and never edits `manifest.json`. Normal tests, static analysis,
LSP, and production CLI commands consume the committed files and must not clone
or contact the network.

The synchronizer:

1. Resolves the repository and exact commit from the manifest (or a caller-
   supplied local checkout) and verifies that checkout `HEAD` equals the full
   pinned SHA. A local checkout with another `HEAD` fails before copying.
2. Builds every project in a unique, invocation-owned staging directory. It
   copies the selected source subtree byte-for-byte, preserving relative names
   and binary contents, then validates `LICENSE`, `SOURCE.md`, and all manifest
   paths.
3. Rejects `.git` entries, symlinks/reparse points, device/FIFO/socket or other
   irregular files, path escapes, missing source projects, and provenance or
   metadata mismatches. Copy order follows sorted relative paths so identical
   pins produce identical trees.
4. Refuses to overwrite a managed destination with uncommitted changes. After
   validation, it publishes the complete managed `projects/third_party` tree
   as one replace operation. Staging and backup directories are unique and
   owned by the invocation; a publish/rename failure restores the prior tree
   and leaves unrelated temporary roots untouched.

The operation is deterministic and idempotent: running it twice at the same
pin yields the same file names, bytes, manifest order, and tree digest, and
removes stale files from a prior managed tree. Pin updates are explicit review
changes to the manifest followed by a synchronization run; branch names and
implicit "latest" updates are not accepted.

## Runner boundary

The native self-corpus runner discovers `example/*` directories that contain
`xlflow.toml` and identifies them as `self/<directory>`. Callers may select a
stable subset by project ID; the default is every discovered project in
lexicographic order. Each project is loaded and analyzed from its own root
through the production `config.Load`, lint, and analyze implementations. The
runner never merges source trees, treats ordinary diagnostics as successful
results, and reports invalid configuration, parser failures, and execution
failures separately from normalized diagnostic records.

The runner is test/developer infrastructure. It reads source files only and
does not modify examples, invoke the corpus synchronizer, fetch upstream
content, open Excel, or require COM/VBE. Third-party manifest profiles and
vendored projects remain governed by the synchronization contract above; a
third-party adapter, golden snapshots, and snapshot-delta reporting are
follow-up work and must not change this sync contract.

## Verification requirements

Manifest tests cover valid v1 data and rejection of unknown fields, unsupported
schema/profile/origin, malformed or abbreviated SHAs, missing values, duplicate
IDs/paths, unsorted projects, path traversal, empty notes, and missing
attribution metadata. Repository checks verify each listed project, license,
and `SOURCE.md` against the manifest.

Synchronization tests use temporary local Git repositories and cover the
correct pinned commit, project boundaries, metadata presence, idempotent tree
digests, stale-file removal, commit mismatch, dirty-tree protection, and
publish failure rollback. The synchronizer additionally rejects missing
projects, symlink/reparse entries, and irregular files during staging. A
second run from the real upstream pin must produce no corpus diff.

Native runner regression tests cover discovery and stable project-ID selection,
independent execution of each project root with its own configuration,
separation of invalid-configuration, parser, and execution failures from
ordinary diagnostics, deterministic diagnostic ordering across repeated runs,
and preservation of partial results when another project fails. The real-world
`self/gen-qrcode` test executes both lint and analyze through the production
implementations, verifies repeatable diagnostics, and compares the project
tree before and after execution to prove that corpus analysis is read-only.
These tests protect the project-isolation and failure-boundary contracts that
unit fixtures alone cannot establish.

The corpus does not change static-analysis semantics, public CLI/API output,
Excel/VBE oracle behavior, or Go dependency-inventory checks. Those suites
remain separate from synchronization and may run fully offline.

## Related

- `docs/adr/ADR-0029-vendored-static-analysis-corpus.md`
- `docs/adr/ADR-0024-shared-static-analysis-rule-registry.md`
- `docs/adr/ADR-0026-local-vbe-oracle-evidence.md`
- Issue #530
- Issue #531
