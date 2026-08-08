# ADR-0031: Deterministic Static-Analysis Diagnostic Snapshots

## Status

Accepted

## Context

ADR-0029 makes the real-world VBA corpus reproducible by pinning and vendoring
the selected upstream projects. ADR-0030 then gives each project an isolated
workspace and remaps diagnostics back to stable project-relative source paths.
Those decisions make analyzer output suitable for regression evidence, but
they do not imply that the projects should produce no findings. Treating every
diagnostic as an error would reject useful parser and analyzer evidence, while
accepting any non-zero output would let analyzer regressions pass unnoticed.

Issue #534 needs a reviewed baseline of the current `lint` and `analyze`
behavior. Diagnostic prose and temporary workspace details are not stable
semantic identity: wording changes should not produce a large unrelated golden
diff, and an invocation-specific temporary path must never become a fixture
contract. At the same time, added and removed diagnostics, position changes,
and duplicate-count changes must remain visible. Baseline generation also needs
to be safe: a normal test must not rewrite its own expected output, and a
failed or partial run must not publish a misleading snapshot.

## Decision

Store one UTF-8 JSONL snapshot per project and analysis surface under
`testdata/static-analysis-corpus/snapshots/<surface>/`. `lint` and `analyze`
each have a file for every analyzed project, using the runner's stable project
ID (`self/<directory>` for native examples and `third_party/<manifest-id>` for
vendored projects) and retaining ID path segments in the snapshot path.
Disabled manifest projects remain documented skips and do not require snapshot
files. A successful surface with no diagnostics is represented by a present
empty file; an absent file is a missing baseline and fails verification.

Each JSONL row is a normalized diagnostic with this exact identity:

```json
{
  "project": "third_party/vba-json",
  "file": "JsonConverter.bas",
  "surface": "analyze",
  "code": "VBA209",
  "severity": "warning",
  "line": 123,
  "column": 5
}
```

`project` and `surface` identify the runner result. `file` is the
slash-separated path relative to the corpus project after source-map remapping;
temporary workspace roots, absolute paths, and generated config paths are
excluded. `line` and `column` are the diagnostic's 1-based start position;
`column` is `0` when no column is available. Rows are sorted by
`project`, `surface`, `file`, `line`, `column`, `code`, and `severity` using a
stable lexical/numeric ordering. Equal rows are not deduplicated: repeated
rows preserve duplicate diagnostic multiplicity. JSONL has no header or blank
lines and uses deterministic JSON serialization.

Message prose, reasons, suggestions, nearby source, and other explanatory
fields are intentionally not part of the snapshot identity. A wording change
therefore does not alter the baseline. If prose becomes a separate contract,
it must receive dedicated tests rather than being added to these golden rows.

The normal corpus test is verify-only. It runs the selected projects through
both production surfaces, normalizes their diagnostics, loads the committed
files, and compares the complete ordered row lists exactly. Additions,
removals, position or identity changes, and duplicate-count changes fail the
test. A diagnostic row records observed analyzer behavior; it is not a claim
that the finding is a true positive or that the upstream project is approved.
Configuration, parser, workspace, cleanup, and execution failures are separate
run failures and fail verification; they are never converted into diagnostic
rows.

Snapshot updates are an explicit developer operation. Run
`task corpus:update-snapshots`, backed by
`XLFLOW_UPDATE_CORPUS_SNAPSHOTS=1`; normal tests never rewrite committed files.
An update must finish every analyzed project and each applicable surface
without any failure before publishing. The updater stages all files and
replaces the snapshot tree atomically; if any project, surface, normalization,
or cleanup step fails, it publishes nothing and leaves all previous files
unchanged. Empty successful snapshots are still written during an explicit
update so a reviewed no-diagnostic result remains stable.

## Consequences

- Real-world analyzer output becomes a deterministic, reviewable regression
  baseline without requiring zero findings from third-party code.
- Stable project IDs and remapped relative paths keep snapshots portable across
  temporary workspace locations and operating systems.
- Added, removed, moved, or duplicated diagnostics fail verification, while
  explanatory wording can evolve without invalidating every baseline row.
- JSONL keeps one finding per diff line and preserves duplicate multiplicity;
  reviewers must still interpret a baseline as observed behavior rather than
  semantic approval.
- Baseline updates are deliberate review changes. A failed run cannot leave a
  partially refreshed tree, but contributors must rerun a complete successful
  update when analyzer behavior intentionally changes.
- Empty snapshot files are meaningful committed artifacts, so cleanup or
  tooling that treats an absent file as an empty result would violate the
  contract.

## Alternatives Considered

1. **Require zero diagnostics from every corpus project** - Rejected because
   third-party projects are evidence for parser/analyzer behavior and may
   contain legitimate findings; this would discard useful coverage.
2. **Snapshot full diagnostic envelopes, including prose** - Rejected because
   message, reason, suggestion, nearby code, and temporary paths are unstable
   and would create noisy, non-semantic diffs.
3. **Hash an unordered or deduplicated diagnostic set** - Rejected because
   ordering would be opaque in reviews and duplicate diagnostics or their
   multiplicity could be lost.
4. **Rewrite snapshots whenever the normal test runs** - Rejected because it
   turns regressions into silent updates and mutates the worktree during test
   verification.
5. **Publish per-project files as each surface finishes** - Rejected because
   a later failure would leave a mixed-time baseline that never represented one
   complete analyzer run.

## Evidence

- Issue #534 acceptance requirements for normalized lint/analyze snapshots,
  deterministic ordering, prose exclusion, explicit updates, and detection of
  added or removed diagnostics.
- Corpus contract and snapshot layout: `docs/specs/static-analysis-corpus.md`.
- Vendored pin and provenance boundary:
  `docs/adr/ADR-0029-vendored-static-analysis-corpus.md`.
- Isolated workspaces, host profiles, project IDs, and source-map remapping:
  `docs/adr/ADR-0030-third-party-corpus-workspaces.md`.
- Runner normalization and failure boundary:
  `internal/staticanalysis/corpus/runner.go`.

## Supersedes

None.

## Superseded by

None.

## Related

- Issue #534
- `docs/specs/static-analysis-corpus.md`
- `docs/adr/ADR-0029-vendored-static-analysis-corpus.md`
- `docs/adr/ADR-0030-third-party-corpus-workspaces.md`
- `docs/adr/ADR-0032-reviewed-static-analysis-corpus-evidence.md`
