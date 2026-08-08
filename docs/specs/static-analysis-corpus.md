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
      },
      "classifications": [{ "path": "examples/analytics/AnalyticsSheet.cls", "kind": "document" }]
    }
  ]
}
```

`notes` is optional; when present it must contain non-whitespace text without
leading or trailing padding. A disabled project must provide `notes` as its
documented skip reason. `classifications` is an optional path-sorted list of
exact project-relative source paths and module kinds (`standard`, `class`,
`form`, or `document`). The default mapping is `.bas` to `standard`, `.cls` to
`class`, and `.frm` to `form`; the adapter never infers document-module
semantics from names or `Attribute VB_*` values. The allowed profiles are
`excel`, `generic-vba`, and `access`. `source.origin` is
currently only `tree-sitter-vba`. `provenance.license` is an SPDX identifier.
The `license_file` and `source_file` values are required relative paths inside
the destination project and must be present in the checked-in tree. `SOURCE.md`
records the original project, repository, license, fixture purpose, and any
normalization or file-removal performed during import.

The loader uses strict JSON decoding. It rejects unknown fields, unsupported
schema/profile/origin values, missing required values, duplicate IDs or
destination/source paths, duplicate provenance paths, malformed repository
names or commit hashes, overlapping project paths, and unsorted project or
classification entries. It also rejects classification traversal, duplicate
paths, unsupported extensions/kinds, and disabled projects without a reason.
IDs are stable and lowercase; projects are ordered lexicographically by ID.
All paths use `/`,
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
content, open Excel, or require COM/VBE.

### Third-party workspace adapter

Enabled manifest projects are analyzed independently. For each project the
adapter creates an invocation-owned temporary child containing
`src/modules`, `src/classes`, `src/forms`, `src/workbook`, and a generated
`xlflow.toml`. Source files retain their project-relative subdirectories and
bytes. The generated config applies the profile policy and uses `frm` as the
UserForm code source. `.cls` document modules are placed under `src/workbook`
only when an exact manifest classification says `document`.

The result project ID is `third_party/<manifest-id>` and the diagnostic file is
the stable project-relative source path. Temporary paths are never emitted in
normalized diagnostics. Missing source-map entries, unsupported layouts,
copy collisions, irregular files, and cleanup failures are reported as
workspace failures; they do not silently fall back to temporary paths.
Disabled projects are skipped with their manifest `notes` reason. Cleanup is
performed after both analysis surfaces, including failure paths, and removes
only the workspace child owned by that invocation.

Profiles are selected in the corpus adapter, not in the shared static-analysis
registry. `excel` keeps the normal rule set. `generic-vba` and `access` omit
the Excel object-model rules `VB002`, `VB003`, `VB027`, `VBA104`, `VBA201`,
`VBA203`, `VBA205`, `VBA211`, `VBA215`, `VBA216`, `VBA217`, `VBA218`, `VBA221`,
`VBA225`, and `VBA226` from corpus evidence. Configurable rules are disabled
in the generated project config; always-on `VBA104` and `VBA211` are filtered
at normalization. This policy affects corpus evidence only and does not alter
production analyzer semantics.

### Deterministic diagnostic snapshots

The corpus records reviewed analyzer behavior as deterministic JSONL snapshots;
it does not require a real-world project to produce zero diagnostics. Snapshot
files live below `testdata/static-analysis-corpus/snapshots/<surface>/`, with
one file per analyzed project (`<project-id>.jsonl`) for each `lint` and
`analyze` surface. Disabled manifest projects remain documented skips and do
not require snapshot files. Project ID path segments are retained, so a
vendored project such as
`third_party/vba-json` is stored at
`snapshots/analyze/third_party/vba-json.jsonl`. Native projects use their
stable `self/<directory>` IDs. A snapshot is UTF-8 JSONL with one normalized
diagnostic object per line, no header or blank lines, and a trailing newline
for every row. A successful surface with no diagnostics still has a present,
zero-byte JSONL file; a missing file is not equivalent to an empty baseline.

The normalized object is the exact snapshot identity:

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

`project` is the stable runner ID, `file` is the slash-separated
project-relative source path after the third-party source-map remap, and
`surface` is `lint` or `analyze`. `line` and `column` are the diagnostic's
1-based start position; `column` is `0` when the normalized diagnostic has no
column. Temporary workspace paths, absolute paths, and generated
configuration paths are never emitted. Multiple diagnostics with the same
identity are retained as repeated rows; comparison therefore preserves their
multiplicity rather than deduplicating them.

Rows are sorted deterministically by `project`, `surface`, `file`, `line`,
`column` (missing columns sort before present columns), `code`, and
`severity`. JSON serialization uses the same field order shown above. Message
prose (`message`, `reason`, `suggestion`, nearby source/code, and other
explanatory fields) is deliberately excluded from identity and JSONL. Wording
changes therefore do not rewrite an otherwise equivalent baseline; prose
contracts, if needed, must be tested separately.

Normal corpus regression is verify-only. It runs every selected project through
both production surfaces, normalizes diagnostics, loads the committed file for
each project/surface, and compares the ordered rows exactly (including duplicate
counts). Added, removed, or moved diagnostics and any identity-field change
fail the test. A diagnostic is observed analyzer output, not an assertion that
the finding is a true positive or semantic approval. Configuration, parser,
workspace, cleanup, or execution failures are separate corpus failures and
also fail verification; they must never be represented as diagnostic rows.

When a snapshot comparison fails, the test renders the structured delta as a
human-readable report. The report starts with `Real-world static-analysis
corpus changed`, then groups rows by diagnostic `code` in lexical order. Each
group header shows `+<added> -<removed>`, followed by every changed row with a
`+` or `-` marker, the stable `project/file` path, the 1-based start position,
and `[surface severity]`. A zero column is rendered as `:line`; a present
column is rendered as `:line:column`. Rows within each direction are ordered
by project, file, line, column, surface, and severity. All rows are retained,
including duplicate identities, so large rule-wide changes remain actionable
after the aggregate count is read.

For example:

```text
Real-world static-analysis corpus changed

VBA206 +0 -1
  - self/example/src/Main.bas:10:2 [analyze warning]

VBA209 +1 -0
  + third_party/vba-web/WebClient.cls:123 [lint warning]
```

Snapshot generation is an explicit developer action, for example
`task corpus:update-snapshots` (which sets
`XLFLOW_UPDATE_CORPUS_SNAPSHOTS=1` for the focused snapshot test). Ordinary
tests never rewrite committed files and do not infer update mode from a diff.
An update must complete every analyzed project and each applicable surface
without any failure before replacing snapshots. The writer stages all output
and publishes it atomically, so a failure leaves every existing snapshot
unchanged; there is no partial update. Empty successful surfaces are written as
empty files during an explicit update so that a reviewed no-diagnostic result
remains a committed baseline. Updating a baseline is a reviewable change to
test data, not an automatic acceptance of newly observed analyzer behavior.

## Operational workflow

The corpus has three deliberately separate evidence layers. Focused unit
fixtures are small, deterministic source snippets owned by the rule tests. They
prove one parser, diagnostic, severity, or LSP contract at a time and should
remain the first place to reproduce a suspected regression. The real-world
corpus runs the production runner against the pinned native and vendored
projects; it exercises project isolation, source mapping, and interactions
between rules that a focused fixture cannot model. The VBE oracle is a local,
developer-only authority for whether a focused VBA case compiles in the real
Excel/VBE environment. It is not a policy or maintainability oracle, and it is
never a replacement for either unit fixtures or the real-world run.

Run the Excel-free contracts before investigating a corpus change. On Windows
use the repository Go wrapper (which selects the supported CGO toolchain); on
other platforms the direct Go command is equivalent:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./internal/staticanalysis/... -count=1
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./internal/oracle -run '^TestCommittedFixtureContractsWithoutExcel$' -count=1
```

For a rule-level change, narrow the focused fixture run to the owning
diagnostic packages before running the corpus-wide checks:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./internal/lint ./internal/analyze -run 'TestLinter|Test.*Analyze' -count=1
```

```bash
rtk go test ./internal/staticanalysis/... -count=1
rtk go test ./internal/oracle -run '^TestCommittedFixtureContractsWithoutExcel$' -count=1
```

```bash
rtk go test ./internal/lint ./internal/analyze -run 'TestLinter|Test.*Analyze' -count=1
```

Then verify the checked-in observed behavior without writing files:

```powershell
rtk task corpus:test
```

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./internal/staticanalysis/corpus -run '^TestRealWorldCorpusSnapshots$' -count=1
```

```bash
rtk go test ./internal/staticanalysis/corpus -run '^TestRealWorldCorpusSnapshots$' -count=1
```

Every snapshot row is an observation of the analyzer version and pinned source
tree at the time it was reviewed. A row is not a claim that the diagnostic is
a true positive, that the source is approved, or that the rule is correct. A
new row therefore requires diagnosis; simply regenerating the file is not a
false-positive fix.

### Pin and vendor changes

Upstream refreshes are explicit data changes. Review the complete manifest
change (including the full 40-character commit, source path, profile,
provenance, and `SOURCE.md`) before running the synchronizer. From the
repository root, use the supported entry point:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\sync-static-analysis-corpus.ps1
```

When the exact pinned commit is already available locally, the same operation
can avoid a network fetch:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\sync-static-analysis-corpus.ps1 -UpstreamCheckout C:\src\tree-sitter-vba
```

The checkout `HEAD` must equal the manifest pin. The synchronizer's clean-tree
check, staging, byte-preserving copy, metadata validation, and atomic publish
are the only supported vendor update path; do not copy files by hand or edit a
snapshot to hide a vendor diff. Inspect the resulting tree and `SOURCE.md`,
run the Excel-free contracts, and run the real-world snapshot verification
before proposing any snapshot update. A pin change and its resulting snapshot
change are separate review concerns and must be explained separately.

### Reading a snapshot delta

The delta report is grouped by diagnostic code and retains every `+` and `-`
row, including duplicate identities. Read it in this order:

1. Confirm that the project/file path and line mapping are stable. A path or
   line movement caused by a vendor refresh is source churn, not automatically
   an analyzer regression.
2. Check the manifest pin and vendored tree digest. If the source did not
   change, reproduce the delta with the focused unit fixture and inspect the
   rule implementation or registry severity.
3. Run the same verification twice. A changing result indicates an ordering,
   parser recovery, temporary-workspace, or other determinism failure; do not
   update the snapshot until it is fixed.
4. For a deliberate analyzer change, land the focused fixture and its review
   first, then perform one explicit snapshot update and include the delta in
   the change description.

An added diagnostic in a real-world project is not evidence that the project is
wrong, and a removed diagnostic is not evidence that the rule is now correct.
Message-only changes are intentionally absent from identity; if message prose
is a contract, test it in the owning focused fixture instead of broadening the
snapshot schema.

### False-positive handling and VBE escalation

When a delta looks like a false positive, first reduce it to a focused fixture
and add an adjacent accepted control. For a compile-equivalent rule, run the
known controls and the focused VBE case sequentially:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\test-vbe-oracle.ps1 --case known-compile-accept --case known-compile-reject --strict
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\test-vbe-oracle.ps1 --case <focused-case-id> --strict
```

Only a confirmed `accepted` or `rejected` result with confirmed Excel cleanup
may be used to bind analyzer behavior. A timeout, unknown modal, failed
Compile invocation, worker/COM failure, or unconfirmed cleanup is an
infrastructure failure: stop, inspect the owned Excel process, and do not
change the analyzer or promote a fixture. Accepted negative controls must
forbid the bound rule on every declared surface; rejected fixtures should list
those controls in `negative_controls` as described by the VBE oracle contract.
Policy and maintainability observations remain unbound when VBE compilation
cannot establish their meaning. Do not silence a suspected false positive by
deleting a row or weakening the real-world snapshot comparison.

### Explicit snapshot updates

After the source, fixture, and (when applicable) VBE evidence are reviewed,
generate snapshots explicitly:

```powershell
rtk task corpus:update-snapshots
```

```bash
rtk go test ./internal/staticanalysis/corpus -run '^TestRealWorldCorpusSnapshots$' -count=1
```

The update command sets `XLFLOW_UPDATE_CORPUS_SNAPSHOTS=1`; it must complete
every selected project and both surfaces before the atomic snapshot tree is
published. Re-run the verify-only command afterwards. A failed update must
leave the previous tree byte-for-byte unchanged, and a zero-byte successful
surface remains a required reviewed artifact.

### Execution-time reference

The following are review observations, not CI thresholds; record actual elapsed
times in a pull request when a run is unusually slow. A Windows run of the
focused corpus test on 2026-08-08 completed in 41.168 s according to Go's
package timer and 43.959 s wall-clock time. The Linux CI step emits the same
project-ordered timing logs plus `corpus_elapsed_seconds` so future regressions
remain visible. Do not treat this single workstation measurement as a fixed
upper bound; record a new observation when analyzer or corpus changes make the
run unusually slow. The VBE oracle is intentionally slower: its default
per-case timeout is 5 min and a multi-case run can take several minutes because
Excel startup and cleanup are serialized.

Measure a local run when timing matters:

```powershell
rtk powershell -NoProfile -Command "Measure-Command { rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./internal/staticanalysis/corpus -run '^TestRealWorldCorpusSnapshots$' -count=1 } | Select-Object TotalSeconds"
```

Do not parallelize VBE cases or infer a hang from the broad `go test ./...`
duration; Excel/COM tests have a materially different runtime profile from
the Excel-free corpus checks.

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

Third-party adapter tests cover manifest classifications, byte-preserving
materialization of all three vendored projects, generated profile configs,
document-module placement, source-map remapping, temporary-path exclusion,
disabled-project skips, partial-result preservation, and cleanup after both
success and failure. A representative third-party project is analyzed through
the production lint/analyze implementations. Snapshot tests additionally cover
the exact JSONL identity and field order, project-relative path remapping,
deterministic ordering across repeated runs, duplicate-row multiplicity,
prose exclusion, empty snapshots, added/removed diagnostic failures, missing
snapshot failures, verify-only behavior, explicit update generation, and the
no-partial-update guarantee when any project or surface fails.

The corpus does not change static-analysis semantics, public CLI/API output,
Excel/VBE oracle behavior, or Go dependency-inventory checks. Those suites
remain separate from synchronization and may run fully offline.

## Related

- `docs/adr/ADR-0029-vendored-static-analysis-corpus.md`
- `docs/adr/ADR-0030-third-party-corpus-workspaces.md`
- `docs/adr/ADR-0031-deterministic-static-analysis-snapshots.md`
- `docs/adr/ADR-0024-shared-static-analysis-rule-registry.md`
- `docs/adr/ADR-0026-local-vbe-oracle-evidence.md`
- Issue #530
- Issue #531
