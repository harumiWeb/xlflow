# Static-analysis corpus

This specification defines the checked-in VBA projects used as static-analysis
fixtures and the developer-only operation that refreshes them. The corpus is
test data, not a runtime dependency of `xlflow`.

## Location and pinned source

The corpus root is `testdata/static-analysis-corpus`. Its manifest is
`manifest.json`; vendored projects live below
`projects/third_party/<project-id>`. Manifest schema version 2 pins the
following 16 independent projects and records expected `.bas`, `.cls`, and
`.frm` source counts for each one:

| ID                     | Profile       | Upstream project path                            | Destination                                 | `.bas` | `.cls` | `.frm` |
| ---------------------- | ------------- | ------------------------------------------------ | ------------------------------------------- | -----: | -----: | -----: |
| `access-examples`      | `access`      | `examples/third_party/Access-examples-master`    | `projects/third_party/access-examples`      |     11 |     22 |      0 |
| `better-access-charts` | `access`      | `examples/third_party/better-access-charts-main` | `projects/third_party/better-access-charts` |      2 |     22 |      0 |
| `iguana-tex`           | `generic-vba` | `examples/third_party/IguanaTex-master`          | `projects/third_party/iguana-tex`           |     10 |      5 |      9 |
| `json`                 | `generic-vba` | `examples/third_party/json-main`                 | `projects/third_party/json`                 |      3 |      1 |      0 |
| `ronecone`             | `excel`       | `examples/third_party/ROneCOne`                  | `projects/third_party/ronecone`             |      0 |      1 |      0 |
| `selenium-vba`         | `generic-vba` | `examples/third_party/SeleniumVBA-main`          | `projects/third_party/selenium-vba`         |     34 |     15 |      0 |
| `std-vba`              | `generic-vba` | `examples/third_party/stdVBA-master`             | `projects/third_party/std-vba`              |     21 |     93 |      0 |
| `vba-cryptography`     | `generic-vba` | `examples/third_party/VBA.Cryptography-main`     | `projects/third_party/vba-cryptography`     |      5 |      0 |      0 |
| `vba-dictionary`       | `generic-vba` | `examples/third_party/VBA-Dictionary-master`     | `projects/third_party/vba-dictionary`       |      1 |      1 |      0 |
| `vba-fast-dictionary`  | `generic-vba` | `examples/third_party/VBA-FastDictionary-master` | `projects/third_party/vba-fast-dictionary`  |      3 |      1 |      0 |
| `vba-fast-json`        | `generic-vba` | `examples/third_party/VBA-FastJSON-master`       | `projects/third_party/vba-fast-json`        |      2 |      0 |      0 |
| `vba-json`             | `generic-vba` | `examples/third_party/VBA-JSON-master`           | `projects/third_party/vba-json`             |      2 |      0 |      0 |
| `vba-memory-tools`     | `generic-vba` | `examples/third_party/VBA-MemoryTools-master`    | `projects/third_party/vba-memory-tools`     |      3 |      1 |      0 |
| `vba-web`              | `excel`       | `examples/third_party/VBA-Web-master`            | `projects/third_party/vba-web`              |     21 |     23 |      0 |
| `wasabi`               | `generic-vba` | `examples/third_party/wasabi-main`               | `projects/third_party/wasabi`               |     22 |      1 |      0 |
| `webxcel`              | `excel`       | `examples/third_party/webxcel-master`            | `projects/third_party/webxcel`              |     11 |     32 |      0 |

All 16 entries are enabled. Synchronization reproduces every manifest entry.
If a future entry is disabled, it remains in the inventory with a documented
reason. The upstream is `harumiWeb/tree-sitter-vba` at the full
40-character commit `2b944e30c7f76dd3e771d02584b80dd6a4733e4d`. A branch, tag,
abbreviated hash, or floating dependency is not a valid pin.

## Manifest schema v2

The JSON document has exactly these top-level fields:

```json
{
  "schema_version": 2,
  "upstream": {
    "repository": "harumiWeb/tree-sitter-vba",
    "commit": "2b944e30c7f76dd3e771d02584b80dd6a4733e4d"
  },
  "projects": [
    {
      "id": "vba-web",
      "path": "projects/third_party/vba-web",
      "profile": "excel",
      "enabled": true,
      "source_counts": { "bas": 21, "cls": 23, "frm": 0 },
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
documented skip reason. `source_counts` is required and records non-negative
`.bas`, `.cls`, and `.frm` counts for the project. Synchronization and
inventory validation compare these counts against the materialized tree.
`classifications` is an optional path-sorted list of
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

Every project keeps its upstream `LICENSE` (or `LICENSE.txt`) and a
repository-local `SOURCE.md`.
The source file is metadata, not an analyzer fixture, and must not be omitted
when the project is refreshed. The imported Access examples intentionally keep
only exported VBA source files and document removed binaries; VBA-Web and
VBA-JSON document that no source modifications were made apart from permitted
line-ending normalization. `THIRD_PARTY_LICENCES.md` lists all 16 projects
separately from the Go module dependency inventory, including the IguanaTex
CC-BY-3.0 attribution.

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

The full real-world snapshot suite is explicitly opt-in. The
`TestRealWorldCorpusSnapshots` test returns without running the corpus unless
`XLFLOW_RUN_REALWORLD_CORPUS=1` is set. This keeps the focused corpus package
tests part of ordinary `go test ./...` runs without repeating the expensive
project-wide lint/analyze pass. The dedicated CI corpus step and the explicit
`task corpus:test` / `task corpus:update-snapshots` entry points set this guard
for their invocation; callers that run the test directly must set the
variable themselves.

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
`VBA225`, `VBA226`, `VBA238`, `VBA242`, and `VBA243` from corpus evidence. For these
generic/access profiles, configurable rules are disabled in the generated
project config; always-on `VBA104` and `VBA211` are
filtered
at normalization. This policy affects corpus evidence only and does not alter
production analyzer semantics.

Excel-profile workspaces explicitly opt in to `VBA242`
(`detect_expensive_full_range_operations`) and `VBA243`
(`detect_value2_performance_opportunities`) so full-range and Value2
performance evidence is captured in the Excel corpus. Generic VBA and Access
profiles disable and exclude `VBA242` and `VBA243` because they do not
establish Excel object-model identity.

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

### Coverage inventory and diagnostic summaries

Manifest `source_counts` provide a deterministic coverage contract independent
of snapshot contents. The inventory helpers scan the materialized
`projects/third_party` tree, compare each project's observed extension counts
with the manifest, and report missing or unexpected project directories. A
clean inventory for the current pin contains 16 projects and 378 VBA source
files: 151 `.bas`, 218 `.cls`, and 9 `.frm`. Profile distribution is 2
`access`, 11 `generic-vba`, and 3 `excel` projects; 16 are enabled and 0 are
disabled. If a future entry is disabled, it remains in the inventory so an
explicit disable cannot hide a coverage change.

`BuildInventory`/`ValidateInventory` are read-only and deterministic; a source
file count mismatch, missing project, stale project directory, symlink, or
irregular file is an error. `SummarizeDiagnostics` groups a runner report by
diagnostic code, surface, and project without deduplicating findings, while
`FormatInventorySummary` and `FormatDiagnosticSummary` provide stable review
logs. These summaries are discovery evidence, not true-positive or precision
claims.

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
`XLFLOW_RUN_REALWORLD_CORPUS=1` and
`XLFLOW_UPDATE_CORPUS_SNAPSHOTS=1` for the focused snapshot test). Ordinary
tests never rewrite committed files and do not infer update mode from a diff.
An update must complete every analyzed project and each applicable surface
without any failure before replacing snapshots. The writer stages all output
and publishes it atomically, so a failure leaves every existing snapshot
unchanged; there is no partial update. Empty successful surfaces are written as
empty files during an explicit update so that a reviewed no-diagnostic result
remains a committed baseline. Updating a baseline is a reviewable change to
test data, not an automatic acceptance of newly observed analyzer behavior.

### Reviewed diagnostic evidence

Deterministic snapshots record observed analyzer output. Reviewed semantic
evidence is a separate UTF-8 JSONL ledger at
`testdata/static-analysis-corpus/reviews/diagnostics.jsonl`. The ledger has no
header or blank lines and uses one object per reviewed diagnostic identity. Its
schema is:

```json
{
  "schema_version": 1,
  "project": "self/gen-qrcode",
  "file": "src/modules/QrCode.bas",
  "classification": "false-positive",
  "diagnostic": {
    "code": "VBA209",
    "severity": "warning",
    "surface": "analyze",
    "range": {
      "start_line": 21,
      "start_column": 5,
      "end_line": 24,
      "end_column": 18
    },
    "count": 1,
    "allowed_occurrences": 2
  },
  "rationale": "The guarded value is initialized on every reachable path.",
  "regression_test": "internal/analyze/qrcode_test.go::TestAnalyzeGuardedQrCodeValue"
}
```

The fields have the following contract:

- `schema_version` is `1`. Unknown versions and unknown fields are rejected.
- `project` is the stable corpus runner ID, and `file` is the slash-separated
  project-relative path after source-map remapping. Absolute and temporary
  workspace paths are invalid.
- `classification` is exactly `true-positive` or `false-positive`. Unreviewed
  is the absence of a matching ledger row, not an assertion serialized into
  the ledger.
- `diagnostic.code` is the canonical registry ID, `severity` is a supported
  registry severity, and `surface` is the exact `lint` or `analyze` surface.
- `diagnostic.range` is required and contains all four normalized coordinates.
  Lines are 1-based. Columns use the corpus diagnostic convention, including
  `0` when no column is available. When the production finding has no end
  position, normalization copies its start line and, for a single-line range,
  its start column before ledger matching. All four coordinates are compared
  exactly; a line-only overlap is not a match.
- `count` is a positive diagnostic multiplicity for the exact identity. It
  makes duplicate findings explicit instead of deduplicating them.
- `allowed_occurrences` is optional and valid only for a false-positive row.
  It records how many separately reviewed diagnostics may continue to share
  the same normalized identity when source/sink detail is not part of the
  corpus contract. Verification fails when the actual count exceeds this
  ceiling. It does not change the false-positive evidence `count` used by
  metrics.

- `rationale` is a non-empty human review explanation.
- A `false-positive` row contains exactly one of `regression_test` or
  `regression_exception`. `regression_test` names the focused regression that
  protects the root-cause fix. `regression_exception` explains why such a
  fixture is impractical. Neither field is valid for a `true-positive` row.

Because the snapshot and shared contract intentionally exclude rule-specific
prose and path witnesses, `allowed_occurrences` cannot distinguish an allowed
occurrence disappearing while the reviewed false positive reappears under the
identical normalized identity. It deterministically guards the observable
multiplicity boundary; a future context-rich identity would be required to
distinguish that swap.

Ledger rows are sorted deterministically by `project`, `file`, diagnostic
`surface`, `code`, `severity`, range coordinates, and `classification`.
Duplicate or conflicting rows for the same exact diagnostic identity are
invalid. The complete identity consists of project, file, surface, code,
severity, and all four range coordinates.

A `true-positive` ledger row is an expected diagnostic contract. Corpus
verification requires the current normalized output to contain that identity
with at least the declared `count`; an unexpected disappearance or reduced
multiplicity fails verification, while additional occurrences remain
unreviewed and visible. A `false-positive` row is a forbidden diagnostic
contract. It remains in the ledger after the analyzer is fixed, and verification
fails if any matching diagnostic reappears. When reviewed diagnostics that are
not false positives share the same exact normalized identity,
`allowed_occurrences` preserves that baseline and verification fails only when
the count grows beyond it.

Every current snapshot diagnostic not consumed by a matching true-positive
row is unreviewed unless it violates a false-positive contract. Unreviewed
diagnostics remain permitted and visible in reports; they are never silently
classified as correct. Snapshot verification still compares all observations,
including unreviewed rows, so this review layer does not weaken the broad
regression baseline.

The committed ledger also supplies deterministic quality metrics. The summary
reports reviewed count, unreviewed count, true-positive count, false-positive
count, and per-rule precision; project/profile coverage may also be reported.
`count` contributes multiplicity to all ledger totals. For a rule with reviewed
evidence, precision is `TP / (TP + FP)`. Unreviewed diagnostics are excluded
from both numerator and denominator. False-positive ledger rows continue to
contribute after remediation, so fixing a diagnostic does not erase reviewed
evidence or inflate precision. Metrics are claims about the committed reviewed
ledger, not estimates for all VBA code.

When the real-world corpus exposes a suspected false positive, use this
workflow:

```text
real-world corpus finds suspicious diagnostic
        ↓
review the exact diagnostic and confirm the false positive
        ↓
reduce the cause to a minimal focused fixture where practical
        ↓
fix the root cause
        ↓
add the focused regression test
        ↓
add or retain the real-world forbidden ledger contract
```

The focused fixture proves the analyzer behavior directly and keeps failures
small and actionable; the real-world forbidden contract proves the original
integration boundary. A `regression_exception` is acceptable only when a
focused fixture is impractical and must record that reason. The broad corpus
case complements, rather than replaces, the focused regression.

The expected/forbidden matcher and diagnostic contract representation are
static-analysis-owned infrastructure that can also be consumed by the local
VBE oracle. This reuse does not make the corpus depend on Excel, and it does not
give VBE compilation universal authority over xlflow diagnostics. VBE
acceptance or rejection is relevant to compile-equivalent rules only.
Maintainability, policy, runtime-risk, portability, and Excel-specific
diagnostics may intentionally report code that compiles successfully; compile
success alone is not evidence that such a warning is a false positive.

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

To select a small review-only batch without running the analyzer, query the
committed snapshots and review ledger by canonical rule ID:

```powershell
rtk task corpus:review-candidates -- VBA225
rtk task corpus:review-candidates -- VBA225 10
```

The optional second argument is a positive display limit and defaults to 20.
The task validates the ledger schema and sources, checks committed
true-positive start-position/multiplicity evidence, consumes the multiplicity
already claimed by matching true-positive rows, and lists the remaining
snapshot rows in deterministic project/file/surface/location order.
False-positive evidence does not consume current rows: any allowed colliding
occurrence remains a candidate for independent review. The reported total
retains duplicate occurrences even when the display is limited. Candidate
locations are start-only snapshot coordinates, so this command is a selection
aid, not a ledger-row generator or an exact forbidden-contract check; obtain
the full normalized range from the analyzer before committing semantic
evidence, and use `task corpus:test` for exact TP/FP verification.

To execute only the projects represented by a bounded candidate set and print
the analyzer's complete normalized ranges, use the developer-only detail task:

```powershell
rtk proxy task corpus:review-details -- --rule VBA225 --limit 20
rtk proxy task corpus:review-details -- --rule VBA225 --project third_party/std-vba
rtk proxy task corpus:review-details -- --rule VBA225 --limit 20 --json
```

`corpus:review-details` joins the start-only committed candidates to a fresh
run and retains duplicate occurrences. Its default output is TSV; `--json`
switches the output to indented JSON. It does not classify diagnostics or
write the review ledger. After inspecting the source, a reviewer may request
schema-valid, canonically sorted JSONL on stdout:

```powershell
rtk proxy task corpus:review-draft -- --rule VBA225 --classification false-positive --rationale "Receiver resolution is incorrect." --regression-test "internal/analyze/analyzer_test.go::TestName"
```

True-positive drafts require `--classification true-positive` and a rationale.
False-positive drafts additionally require exactly one of `--regression-test`
or `--regression-exception`. The task never edits
`reviews/diagnostics.jsonl`; the reviewer must inspect and merge its stdout.
The outer `rtk proxy` is required for the detail and draft tasks because their
TSV/JSON/JSONL stdout is an artifact and must not be summarized or truncated.

During implementation, a read-only focused comparison can select a stable
project ID, a rule ID, or both:

```powershell
rtk task corpus:test-focused -- --project third_party/std-vba --rule VBA225
rtk task corpus:test-focused -- --rule VBA225
```

Rule-only verification still executes every project so a new diagnostic in a
previously empty project cannot be hidden. Focused verification checks the
selected snapshot and exact TP/FP contracts, cannot update snapshots, and is
only a fast development loop. `corpus:test` remains the required full gate.

Then verify the checked-in observed behavior without writing files:

```powershell
rtk task corpus:test
```

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -Command '$env:XLFLOW_RUN_REALWORLD_CORPUS="1"; & ".\scripts\dev\go.ps1" test ./internal/staticanalysis/corpus -run "^TestRealWorldCorpusSnapshots$" -count=1'
```

```bash
XLFLOW_RUN_REALWORLD_CORPUS=1 rtk go test ./internal/staticanalysis/corpus -run '^TestRealWorldCorpusSnapshots$' -count=1
```

The direct commands above are opt-in and must include
`XLFLOW_RUN_REALWORLD_CORPUS=1`. Prefer `rtk task corpus:test`, which sets the
guard and the verify-only snapshot mode together. The ordinary package test
command intentionally skips the full suite when the guard is absent.

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

An optimization-induced diagnostic increase follows the same review boundary.
Do not update the snapshot to accept the new row. Re-run the verify-only
comparison twice, reduce the delta to a focused fixture, and classify it as a
true-positive contract change, a false positive, or nondeterminism. A false
positive requires a focused regression and a root-cause fix in the analyzer;
for a compile-equivalent rule, bind the fix to sequential VBE-oracle evidence
and retain the corresponding forbidden negative control. Only after the
focused contract is stable may a reviewed snapshot update be considered.

### Explicit snapshot updates

After the source, fixture, and (when applicable) VBE evidence are reviewed,
generate snapshots explicitly:

```powershell
rtk task corpus:update-snapshots
```

```bash
XLFLOW_RUN_REALWORLD_CORPUS=1 rtk go test ./internal/staticanalysis/corpus -run '^TestRealWorldCorpusSnapshots$' -count=1
```

The update command sets `XLFLOW_RUN_REALWORLD_CORPUS=1` and
`XLFLOW_UPDATE_CORPUS_SNAPSHOTS=1`; it must complete
every selected project and both surfaces before the atomic snapshot tree is
published. Re-run the verify-only command afterwards. A failed update must
leave the previous tree byte-for-byte unchanged, and a zero-byte successful
surface remains a required reviewed artifact.

### Execution-time reference

The following are review observations, not CI thresholds; record actual elapsed
times in a pull request when a run is unusually slow. After all 16 projects
were enabled on 2026-08-08, the snapshot update completed in 202.028 s by the
package timer. Three verify-only Windows runs completed in 201.156 s, 196.764 s,
and 200.100 s wall-clock time; the latter reported a 197.263 s package timer.
Its newly enabled project timings were 1m22.250s for `std-vba`, 8.209s for
`vba-fast-dictionary`, and 11.002s for `vba-fast-json`. These results include
both lint and analyze for each isolated project and remain below the 10-minute
CI budget. The Linux CI step emits the same project-ordered
timing logs plus `corpus_elapsed_seconds` so future regressions remain visible.
Do not treat this single workstation measurement as a fixed upper bound; record
a new observation when analyzer or corpus changes make the run unusually slow.
The VBE oracle is intentionally slower: its default
per-case timeout is 5 min and a multi-case run can take several minutes because
Excel startup and cleanup are serialized.

Measure a local run when timing matters:

```powershell
rtk powershell -NoProfile -Command '$env:XLFLOW_RUN_REALWORLD_CORPUS="1"; Measure-Command { & ".\scripts\dev\go.ps1" test ./internal/staticanalysis/corpus -run "^TestRealWorldCorpusSnapshots$" -count=1 } | Select-Object TotalSeconds'
```

### Developer-only corpus profiling

The slowest projects have a focused benchmark that is separate from snapshot
and review evaluation. It materializes each selected project in its own
temporary workspace and offers two stages: `full-pipeline` runs materialization,
configuration, lint, analyze, normalization, and cleanup; `analyze-only`
materializes and loads configuration outside the timer, then measures the
production analyzer. Benchmarks are opt-in Go test binaries: ordinary tests,
the production CLI/API, corpus CI, and the Excel/VBE oracle never invoke them.

Run the analyzer-only measurement twice for both hotspot projects:

```powershell
rtk task bench:corpus
```

Use a leaf benchmark name to isolate one project and stage. `-count=2` is the
minimum timing run for variability; profiles use `-count=1` so output files are
not overwritten by repeated test processes:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -Command '$cpu = Join-Path $env:TEMP "xlflow-std-vba.cpu.pprof"; $mem = Join-Path $env:TEMP "xlflow-std-vba.mem.pprof"; $bin = Join-Path $env:TEMP "xlflow-std-vba-corpus.test.exe"; $args = @("test", "./internal/staticanalysis/corpus", "-run", "^$", "-bench", "^BenchmarkRealWorldCorpus/std-vba/analyze-only$", "-benchtime=1x", "-count=1", "-cpuprofile=$cpu", "-memprofile=$mem", "-o=$bin"); & ".\scripts\dev\go.ps1" @args'
rtk powershell -NoProfile -ExecutionPolicy Bypass -Command '$cpu = Join-Path $env:TEMP "xlflow-ronecone.cpu.pprof"; $mem = Join-Path $env:TEMP "xlflow-ronecone.mem.pprof"; $bin = Join-Path $env:TEMP "xlflow-ronecone-corpus.test.exe"; $args = @("test", "./internal/staticanalysis/corpus", "-run", "^$", "-bench", "^BenchmarkRealWorldCorpus/ronecone/analyze-only$", "-benchtime=1x", "-count=1", "-cpuprofile=$cpu", "-memprofile=$mem", "-o=$bin"); & ".\scripts\dev\go.ps1" @args'
```

The 2026-08-08 Windows CPU profile attributed 112.24 s cumulative CPU to
procedure data-flow analysis, including 102.96 s in state joins. Repeated
serialization of propagation paths accounted for 82.45 s in `pathKey`, with
41.42 s below `fmt.Fprintf` and 23.77 s below `strings.ToLower`; GC draining
accounted for 133.68 s across the full profiled corpus. The optimized comparator
streams the exact former canonical byte sequence, retains lowercase label and
decimal lexical ordering, and allocates nothing for ASCII path comparisons.
This preserves representative paths and diagnostic behavior while removing the
dominant temporary strings.

The project-specific heap profiles also separated the two slow projects. After
the path comparison change, `ronecone` still allocated 20.89 GB through
`ProjectSummary.All` while resolving call candidates: every lookup defensively
cloned every procedure and evidence slice before selecting one summary. The
replacement maintains a declaration-line index, rechecks the original
slash-normalized `EqualFold` file/name/kind and exact-line contract, preserves
the first deterministic match, and defensively clones only that match.

The focused `LibJSON.ParseChars` benchmark set a target of at least 20% lower
median runtime and 50% lower allocated bytes per operation. On the same
Windows machine, the pre-change five-run baseline was 1.190-1.266 s,
1,003.4-1,003.8 MB, and 20.681-20.682 million allocations per operation. Three
post-change runs were 0.919-0.934 s, 447.6-448.1 MB, and 5.576 million
allocations per operation. Median runtime fell from 1.197 s to 0.924 s (22.8%),
allocated bytes fell by approximately 55.3%, and allocation count fell by
approximately 73.0%, meeting the recorded target.

The final project-isolated `analyze-only` benchmark produced two post-change
samples of 62.869 s and 62.980 s for `std-vba` (0.2% spread), 19.741 GB
allocated per operation, and 2,792 stable findings. For `ronecone`, the
candidate-lookup baseline was 46.014 s and
47.769 s with 43.214 GB allocated per operation; post-index samples were
36.156 s and 36.664 s (1.4% spread), 20.617 GB per operation, and the same 808
findings. Its two-sample midpoint runtime fell by 22.4% and allocated bytes by
52.3%. These numbers are developer observations rather than fixed thresholds;
compare future changes using the same command, machine, power state, and Go
toolchain.

After rebasing onto the context-cancellation analyzer change, two consecutive
final verify-only corpus runs completed in 171.5 s and 170.2 s wall-clock time
with no snapshot, diagnostic, severity, range, multiplicity, or ordering delta.
Compared with the earlier same-machine 196.764-201.156 s reference, the
midpoint fell by approximately 14.1%. `std-vba` completed in 68.276 s and
68.474 s, down 16.9% from the earlier 82.250 s observation; `ronecone`
completed in 38.558 s and 38.635 s, down approximately 16.1% from the roughly
46 s issue baseline. The focused benchmark target remains the hardware-local
optimization guard because it isolates this change from independent analyzer
instrumentation overhead.

Do not parallelize VBE cases or infer a hang from the broad `go test ./...`
duration; Excel/COM tests have a materially different runtime profile from
the Excel-free corpus checks.

### Batch analyzer stage profiling and large single-module benchmarks

Issue #671 adds a separate profiling path for unusually large VBA modules. The
opt-in `xlflow analyze --performance-log` flag emits canonical stage timings,
result counts, outcomes, and workload counters to stderr while leaving analyzer
findings and JSON stdout unchanged. The records are intended for same-machine
comparisons and are not CI timing thresholds. Use `outcome` to distinguish
successful, failed, and canceled stages.

The aggregate counters describe total project work, including files,
procedures, statements, expressions, call sites, CFG blocks/edges, and project
symbols, lines, and module declarations. Subsystem worklist counters may also
be present when object or ByRef analysis runs. The maximum counters describe
the largest single-file or single-procedure dimensions: `max_lines_per_file`,
`max_procedures_per_file`, `max_calls_per_file`,
`max_statements_per_procedure`, `max_cfg_blocks_per_procedure`, and
`max_cfg_edges_per_procedure`. These counters are workload measurements, not
diagnostics. The canonical stage labels include `parse`, `procedure_ir`, `cfg`,
`effect_summaries`, `project_context`, `project_wide_diagnostics`,
`procedure_local_diagnostics`, and `suppression_and_finalize`.

For issue #694, `procedure_local_diagnostics` remains the parent stage and is
also attributed to these aggregate semantic-domain stages:
`procedure_local/source_scan`, `procedure_local/runtime`,
`procedure_local/array`, `procedure_local/object`,
`procedure_local/dictionary`, `procedure_local/error`,
`procedure_local/dataflow`, `procedure_local/resource`,
`procedure_local/excel`, `procedure_local/application_state`, and
`procedure_local/other`. The records are aggregate observations rather than
one timer per procedure. Procedure work may run in parallel, so a domain's
elapsed time is cumulative worker time, not an additive wall-time partition;
domain values can exceed the parent stage's elapsed time. Typed Excel, object
summary/entry state, and project-context index work remain in their existing
top-level stages.

Domain profiling also reports work counters when the corresponding work runs.
Candidate counters are `runtime_candidate_procedures`,
`array_candidate_procedures`, `object_candidate_procedures`,
`dictionary_candidate_procedures`, `error_candidate_procedures`,
`dataflow_candidate_procedures`, `resource_candidate_procedures`,
`excel_candidate_procedures`, and `application_state_candidate_procedures`.
Traversal counters are `source_line_scans`, `runtime_cfg_walks`,
`array_cfg_walks`, `dictionary_cfg_walks`, `error_cfg_walks`,
`dataflow_cfg_walks`, `resource_cfg_walks`, and `excel_cfg_walks`.
`semantic_kernel_runs` counts valid semantic-domain kernel invocations.
Candidates are procedures that pass the relevant rule gate and are actually
analyzed; traversal counters count started source/CFG traversals. These are
workload counters and never counts of emitted findings.

Array-domain profiles additionally report `array_kernel_runs`,
`array_cfg_walks`, and `array_projection_runs`. The first counts one canonical
array semantic-result materialization per applicable procedure revision, the
second counts started array fixed-point walks, and the third counts enabled
and applicable array projectors. Enabling multiple core array diagnostics must
leave the kernel and main CFG-walk counts at one per applicable procedure;
`VBA241` is a shared-fact projection, while an applicable `VBA226` secondary
shape pass is recorded as an explicit additional walk. These counters measure
work, not findings, and are compared alongside `ns/op`, `B/op`, and
`allocs/op`.
The CLI emits stage records with `operation="analyze/stage"` and counter
records with `operation="analyze/counter"`; both remain stderr-only.

Run the deterministic synthetic benchmark and the real-world corpus hotspot
benchmark with the repository tasks:

```powershell
rtk task bench:analyze
rtk task bench:analyze-single-module
rtk task bench:corpus
```

`bench:analyze` retains the existing multi-module and object-worklist baselines.
`bench:analyze-single-module` keeps fixture generation outside the timed region
and covers single-module scales around 100, 500, 1,000, and 2,000 procedures,
including large call graphs, declaration sets, and CFGs. The 2,000-procedure
domain-profiling matrix retains these four workload shapes:

| benchmark                        | workload shape         |
| -------------------------------- | ---------------------- |
| `independent/2000-procedures`    | independent procedures |
| `declarations/2000-declarations` | declaration-heavy      |
| `chain/2000-procedures`          | call-heavy             |
| `cfg-independent/2000-branches`  | CFG-heavy              |

`bench:corpus` runs the checked-in
`std-vba` and `ronecone` analyze-only sub-benchmarks. These tasks use
`-benchmem -benchtime=1x`; on Windows they run through `scripts/dev/go.ps1` to
keep CGO and tree-sitter toolchain selection consistent. Benchmark output should
retain wall time (`ns/op`), allocations, findings, workload dimensions, and the
`stage_*` / `counter_*` metrics derived from the attached analysis recorder;
these are Go benchmark metrics, not stderr records from
`analyze --performance-log`.

For issue #696, the benchmark matrix also covers no-array procedures, simple
dynamic arrays, ReDim-heavy procedures, branch-heavy array lifecycles,
multidimensional arrays, ByRef array flows, large CFGs, and a workload with
all compatible array projectors enabled. Record the array kernel, walk, and
projection counters with the standard wall-time and allocation metrics. A
repeated run must retain deterministic findings and snapshots; a performance
comparison must show reduced array walks or allocations on array-heavy and
ROneCOne workloads without multiplying the main fixed-point work when another
projector is enabled.

For issue #675, the benchmark matrix must also include one-file workloads with
100, 500, 1,000, and 2,000 procedures and a many-file workload. Where the
host permits, compare `-cpu=1,2,4,8` (equivalent `GOMAXPROCS` settings) and
record wall time, allocations, finding counts, and stable JSON equality across
repeated runs. The 100-
procedure and many-file cases guard the ordinary file-level fast path; the
1,000- and 2,000-procedure cases measure the intended intra-file speedup. These
measurements are hardware-local evidence, not CI timing thresholds, and the
benchmark must not run from ordinary `analyze`, corpus snapshot, or VBE-oracle
workflows.

For issue #677, the same Windows amd64 machine (12th Gen Intel(R) Core(TM)
i7-12700, Go 1.26.6 toolchain) was used with baseline HEAD
`245d4ad0792bbc49c2769e3cbd7f0156dc9a7c33` and the implementation worktree,
using `BenchmarkSingleModuleSynthetic/independent/2000-procedures` and
`BenchmarkSingleModuleSynthetic/declarations/2000-declarations`,
`-benchmem -benchtime=1x -count=5`. The independent trial values
(ns/op; B/op; allocs/op) were, before:
`15,082,456,700/10,595,106,376/175,671,258`,
`14,070,609,300/10,595,213,184/175,673,277`,
`14,874,702,400/10,594,734,880/175,680,863`,
`15,055,094,800/10,595,629,288/175,672,346`, and
`13,300,470,200/10,595,969,152/175,667,154`; after:
`1,801,398,600/1,453,074,600/11,277,336`,
`1,812,785,500/1,452,991,792/11,281,779`,
`1,859,160,700/1,453,007,720/11,287,788`,
`2,106,054,200/1,453,114,664/11,283,437`, and
`1,914,653,900/1,452,904,888/11,279,283`. Medians changed from
14,874,702,400 ns/op, 10,595,213,184 B/op, and 175,672,346 allocs/op to
1,859,160,700 ns/op, 1,453,007,720 B/op, and 11,281,779 allocs/op: 8.00x
speedup, with 87.5% less time, 86.3% fewer allocated bytes, and 93.6% fewer
allocations. The recorder reported `module_fact_builds=1` and
`procedure_fact_builds=2000` per operation.

The declaration-heavy trial values were, before:
`2,794,628,900/3,600,116,176/1,065,219`,
`3,083,994,200/3,601,708,584/1,064,689`,
`3,128,503,500/3,601,308,800/1,065,147`,
`2,532,879,000/3,601,569,448/1,064,520`, and
`2,152,009,600/3,601,638,472/1,064,243`; after:
`1,770,364,600/3,580,182,600/1,044,882`,
`1,615,096,000/3,577,576,544/1,044,232`,
`1,481,524,900/3,578,309,560/1,044,122`,
`1,601,773,700/3,578,604,872/1,044,057`, and
`1,508,833,100/3,577,864,840/1,045,309`. Medians changed from
2,794,628,900 ns/op, 3,601,569,448 B/op, and 1,064,689 allocs/op to
1,601,773,700 ns/op, 3,578,309,560 B/op, and 1,044,232 allocs/op: 1.74x
speedup, with 42.7% less time, 0.65% fewer allocated bytes, and 1.92% fewer
allocations. Timing is observational and is not a CI assertion. Retain the
complete trial outputs, SHA, command, and power-state notes with performance
investigations; do not treat these measurements as a semantic or diagnostic
contract.

Issue #674 adds an effects-focused scalability measurement to this procedure.
The benchmark generator must construct the IR, CFG, and project-local call
graph before the timer starts, then measure only `effects.Build` (or its
equivalent package benchmark). Run long chains, wide fan-in, wide fan-out,
bounded-dense project-local graphs, large uncertainty sets, and effect-heavy
callees at approximately 500, 1,000, and 2,000 procedures. Each run records
wall time (`ns/op`), `B/op`, `allocs/op`, maximum propagated facts per
procedure, total propagated facts, and worklist evaluations. Keep the Go
toolchain, machine, power state, benchmark filter, `-benchmem`,
`-benchtime=1x`, and sample count fixed for before/after comparisons; on
Windows invoke Go through `scripts/dev/go.ps1`.

The baseline is captured before changing the propagation algorithm. Report
the median and spread for at least five before/after samples for the focused
effects cases. A successful optimization should reduce the provenance-heavy
median wall time by at least 20% and allocated bytes by at least 50% without
increasing allocation count; record the observed result rather than treating
these targets as a CI timing threshold. Re-run the existing
`BenchmarkSingleModuleSynthetic` scales and the `std-vba`/`ronecone`
`analyze-only` leaf benchmarks to determine whether `effect_summaries` is a
confirmed end-to-end hotspot. If it is, retain CPU and heap profiles for the
same leaf benchmark; if it is not, record that result and do not attribute the
workload-wide improvement to effects.

The 2026-08-20 Windows observation used the i7-12700 host, the repository Go
wrapper, the same benchmark generator, and five samples for the effects
comparison. The legacy rows were collected by temporarily selecting the
pre-#674 eager propagation function; that switch was not retained in the
source tree. Values below are medians from `-benchtime=1x -benchmem`:

| workload             | eager `ns/op` | bounded `ns/op` | wall change | eager `B/op` | bounded `B/op` | bytes change | eager `allocs/op` | bounded `allocs/op` |
| -------------------- | ------------: | --------------: | ----------: | -----------: | -------------: | -----------: | ----------------: | ------------------: |
| effect-heavy / 500   |     116.82 ms |        13.48 ms |      -88.5% |     24.65 MB |        8.01 MB |       -67.5% |              364k |               93.5k |
| effect-heavy / 1,000 |     206.30 ms |        27.85 ms |      -86.5% |     51.18 MB |       17.74 MB |       -65.3% |              737k |              196.9k |
| effect-heavy / 2,000 |     227.69 ms |        57.67 ms |      -74.7% |    104.31 MB |       37.37 MB |       -64.2% |            1.492m |                409k |

The bounded runs reported maximum semantic propagated facts of 14 and
worklist evaluations of 999, 1,999, and 3,999 for the three sizes. Chain and
dense workloads also reduced wall time (about 66--89% and 59--75%
respectively), but their byte reductions were below the 50% target; the target
is therefore claimed only for the provenance-heavy effect-heavy cases. The
optimized ROneCOne `analyze-only` leaf completed in 108.35 s and 126.28 s; its
`effect_summaries` stage was 0.308 s and 0.373 s (roughly 0.3% of total), so it
was not a confirmed end-to-end hotspot. The legacy ROneCOne attempt terminated
before producing a benchmark record in this workspace; no ROneCOne speedup is
claimed from that incomplete before run.

ROneCOne profiling is developer-only and must be selected explicitly as a leaf
benchmark; it is Excel/COM-free and is not part of ordinary `go test ./...`:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File '.\scripts\dev\go.ps1' test ./internal/staticanalysis/corpus -run '^$' -bench '^BenchmarkRealWorldCorpus/ronecone/analyze-only$' -benchmem -benchtime=1x -count 5
```

For a reproducible CPU and allocation profile, use `-count=1` and keep the
benchmark binary and profiles outside the repository. The benchmark
materializes and loads the fixture before the timed `analyze-only` region.
Go's process-wide profile may still observe the small amount of benchmark
setup, but the leaf filter keeps the captured workload focused on analyzer
work:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -Command '$cpu = Join-Path $env:TEMP "xlflow-ronecone.cpu.pprof"; $mem = Join-Path $env:TEMP "xlflow-ronecone.mem.pprof"; $bin = Join-Path $env:TEMP "xlflow-ronecone-corpus.test.exe"; $args = @("test", "./internal/staticanalysis/corpus", "-run", "^$", "-bench", "^BenchmarkRealWorldCorpus/ronecone/analyze-only$", "-benchtime=1x", "-count=1", "-cpuprofile=$cpu", "-memprofile=$mem", "-o=$bin"); & ".\scripts\dev\go.ps1" @args'
rtk go tool pprof -top "$env:TEMP\xlflow-ronecone-corpus.test.exe" "$env:TEMP\xlflow-ronecone.cpu.pprof"
rtk go tool pprof -sample_index=alloc_space -top "$env:TEMP\xlflow-ronecone-corpus.test.exe" "$env:TEMP\xlflow-ronecone.mem.pprof"
rtk go tool pprof -sample_index=alloc_objects -top "$env:TEMP\xlflow-ronecone-corpus.test.exe" "$env:TEMP\xlflow-ronecone.mem.pprof"
rtk go tool pprof -sample_index=inuse_space -top "$env:TEMP\xlflow-ronecone-corpus.test.exe" "$env:TEMP\xlflow-ronecone.mem.pprof"
```

Add `-cum` to the first `pprof` view when cumulative CPU consumers are needed;
without it the view is ordered by flat CPU. The next two report top
allocation-space and allocation-object consumers, and the final view reports
heap in use when that profile sample is available. Keep the complete command
output with the benchmark record.

On Linux or macOS, use this directly executable POSIX equivalent. It writes a
native test binary (without the Windows `.exe` suffix) and keeps profiles in a
temporary directory:

```sh
profile_dir="${TMPDIR:-/tmp}/xlflow-ronecone-694"
rtk mkdir -p "$profile_dir"
rtk go test ./internal/staticanalysis/corpus \
  -run '^$' \
  -bench '^BenchmarkRealWorldCorpus/ronecone/analyze-only$' \
  -benchtime=1x -count=1 \
  -cpuprofile="$profile_dir/ronecone.cpu.pprof" \
  -memprofile="$profile_dir/ronecone.mem.pprof" \
  -o="$profile_dir/ronecone-corpus.test"
rtk go tool pprof -top "$profile_dir/ronecone-corpus.test" "$profile_dir/ronecone.cpu.pprof"
rtk go tool pprof -sample_index=alloc_space -top "$profile_dir/ronecone-corpus.test" "$profile_dir/ronecone.mem.pprof"
rtk go tool pprof -sample_index=alloc_objects -top "$profile_dir/ronecone-corpus.test" "$profile_dir/ronecone.mem.pprof"
rtk go tool pprof -sample_index=inuse_space -top "$profile_dir/ronecone-corpus.test" "$profile_dir/ronecone.mem.pprof"
```

Keep the Go version, machine, power state, benchmark filter, and sample count
constant for before/after comparisons, and report unusually slow runs with the
complete command and environment. Benchmark generation is test-only
infrastructure and must not run from ordinary `analyze`, `check`, corpus
snapshot, or Excel/VBE oracle workflows.

### Issue #694 baseline record

The post-wave-2 observations above remain historical reference values and must
not be overwritten. After capturing the first rule-domain profiles, append one
completed row per ROneCOne or synthetic workload to the following record shape.
Record the repository SHA, machine/toolchain and power-state notes, complete
command, wall time, `B/op`, `allocs/op`, top CPU consumers, top alloc-space
consumers, and the semantic-domain timings and work counters from the same run.

The initial #694 capture was made on 2026-08-22 from base SHA
`6858757ea938d60e123cbadd2e9facf1f34e4217` plus the #694 working-tree change,
using Go 1.26.6 on Windows 11 Home 10.0.22631, a Thirdwave XA7C-R38 with an
Intel Core i7-12700 (12 cores / 20 logical processors), and the normal
plugged-in developer power state. Five `-benchtime=1x` samples were taken
serially.

| workload                          | `ns/op` samples                                                                |   median | min--max spread |  median `B/op` | median `allocs/op` |
| --------------------------------- | ------------------------------------------------------------------------------ | -------: | --------------: | -------------: | -----------------: |
| ROneCOne analyze-only             | 23,815,264,100; 24,193,761,600; 28,179,640,900; 26,682,322,000; 24,197,555,700 | 24.198 s |         4.365 s | 27,319,440,464 |        178,295,249 |
| independent / 2,000 procedures    | 1,237,728,400; 1,234,423,800; 1,236,714,300; 1,231,832,700; 2,670,839,000      |  1.237 s |         1.439 s |  1,469,165,568 |         11,450,545 |
| declarations / 2,000 declarations | 1,673,953,300; 1,540,513,700; 1,431,431,000; 1,462,106,900; 1,617,206,000      |  1.541 s |         0.243 s |  3,578,884,504 |          1,044,303 |
| chain / 2,000 procedures          | 963,067,500; 864,272,300; 849,723,600; 876,187,500; 856,400,900                |  0.864 s |         0.113 s |  1,501,293,120 |         11,974,108 |

The `cfg-independent/2000-branches` fixture remains contract-tested and
benchmark-selectable, but its five-sample capture did not complete on this
machine: the pre-instrumentation five-run command failed after about 368 s and
a single run exceeded ten minutes. No partial timing is presented as a
baseline. This is a workload result, not a disabled-profiling regression.

For ROneCOne, the median parent `procedure_local_diagnostics` wall time was
3.456 s. Domain values below are median accumulated worker execution times;
they overlap under procedure parallelism and therefore must not be added or
compared as a partition of the parent wall time. In addition,
`procedure_local/source_scan` contains the source-line loop and therefore
includes elapsed time and result counts from nested Excel, object, and array
measurements; those nested values must not be added to source-scan values.

| semantic domain         | median accumulated time |
| ----------------------- | ----------------------: |
| source scan             |                 1.754 s |
| runtime                 |                 4.964 s |
| array                   |                10.001 s |
| object                  |                 1.585 s |
| dictionary / collection |                 3.011 s |
| error                   |                 0.109 s |
| dataflow                |                28.564 s |
| resource                |                 0.014 s |
| Excel-specific          |                 4.604 s |
| application state       |                 0.284 s |
| other                   |                 0.002 s |

The work counters were stable across the five samples: 1,565 procedures for
each candidate domain; 24,303 source lines scanned; 134,289 semantic kernel
runs; 1,565 runtime, dictionary, dataflow, and resource CFG walks; 6,260 array
CFG walks; 4,695 error CFG walks; and 3,130 Excel CFG walks. The object
candidate counter was 1,565 and has no CFG-walk counter by design.

### Applicability-planner verification record

Issue #695 adds immutable procedure feature summaries and applicability-based
skipping for expensive semantic domains. Verify the optimization against the
same conservative analyzer contract used for corpus snapshots; do not accept
planner counters as evidence of a diagnostic change by themselves.

For every implementation change, run the focused planner matrix and the
deterministic corpus snapshot verification. Compare findings, ranges,
severity, multiplicity, ordering, suppression, and exit status with a
conservative-all reference path or equivalent regression fixture. Include
recovered/incomplete procedures, ambiguous or dynamic calls, and incomplete
project/type facts so unknown applicability is demonstrated as planned rather
than skipped. The batch serial/parallel and realtime/LSP paths should use the
same applicability decisions where both surfaces support the rule.

Record `planned_<domain>_runs` and `skipped_<domain>_runs` from the same
`analyze --performance-log` or benchmark run as domain timings. A skipped run
must be a proven-absence decision; unknown features are planned. The counters
are observations of planner work, not findings, and must not be used to
rewrite a snapshot without a separate semantic review.

Performance coverage must include scalar-only, array-heavy, mixed-domain, and
recovered/incomplete synthetic procedures plus ROneCOne. Use the existing
`bench:analyze-single-module` and `bench:corpus` workflows, keep fixture
materialization outside the timed analyzer region, and take five serial
`-benchtime=1x` samples for ROneCOne when collecting a before/after record.
Report planner construction cost, avoided semantic-domain work, wall time,
`B/op`, `allocs/op`, domain timings, planned/skipped counters, and CFG/kernel
work counts. Keep the Go version, machine, power state, SHA, benchmark filter,
and command constant for comparison. Small-module overhead is an acceptance
check, not a fixed CI timing threshold.

The #695 capture was made on 2026-08-22 from base SHA `aa2cddd9` plus the #695
working-tree change. It used the same Thirdwave XA7C-R38, Intel Core i7-12700,
Windows 11 Home 10.0.22631, Go 1.26.6, and normal plugged-in developer power
state as the #694 baseline. Both comparisons used five serial
`-benchtime=1x` samples.

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./internal/staticanalysis/corpus -run '^$' -bench '^BenchmarkRealWorldCorpus/ronecone/analyze-only$' -benchmem -benchtime=1x -count=5
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./internal/analyze -run '^TestSingleModuleBenchmarkFixtureScale$' -bench '^BenchmarkSingleModuleSynthetic/independent/2000-procedures$' -benchmem -benchtime=1x -count=5
```

| workload                       | `ns/op` samples                                                                |   median | change from #694 |  median `B/op` | median `allocs/op` |
| ------------------------------ | ------------------------------------------------------------------------------ | -------: | ---------------: | -------------: | -----------------: |
| ROneCOne analyze-only          | 19,613,948,300; 19,543,799,200; 19,563,849,500; 19,571,910,100; 19,609,888,100 | 19.572 s |           -19.1% | 27,164,873,376 |        179,448,654 |
| independent / 2,000 procedures | 1,222,285,600; 1,185,823,200; 1,164,685,300; 1,164,679,600; 1,166,594,100      |  1.167 s |            -5.7% |  1,489,784,960 |         11,649,544 |

For ROneCOne, planner decisions were stable across all samples: application
state 1,365 planned / 200 skipped, dataflow 1,382 / 183, Excel 1,346 / 219,
and resource 1,365 / 200. Array, dictionary, error, object, and runtime remained
conservatively planned for all 1,565 procedures because call/effect facts did
not prove those domains absent. The resulting dataflow CFG walks fell from
1,565 to 1,382, resource CFG walks from 1,565 to 1,365, Excel CFG walks from
3,130 to 2,692, and semantic kernel runs from 134,289 to 132,247. Median
accumulated domain times were 1.181 s source scan, 2.915 s runtime, 6.428 s
array, 0.910 s object, 1.866 s dictionary/collection, 0.071 s error, 17.330 s
dataflow, 0.009 s resource, 2.891 s Excel, and 0.096 s
application state. These worker times overlap and are not additive.

The independent synthetic workload remained conservative-all because its
generated control-flow facts do not prove domain absence. Even with planner
construction and counter reporting, its median wall time did not regress;
median bytes and allocations increased by 1.4% and 1.7%, respectively, within
the existing 2% observational criterion. The microbenchmark separately covered
scalar-only, array-heavy, mixed-domain, and recovered procedures with
`-benchtime=100x -count=3` (100 `b.Loop()` iterations per sample). Planner
construction used 0 B/op and 0 allocs/op; the three `plan` samples ranged from
662--957 ns/op (scalar-only), 645--685 ns/op (array-heavy), 689--720 ns/op
(mixed-domain), and 651--664 ns/op (recovered). The corresponding
`feature-facts` samples ranged from 1,067--1,306 ns/op (scalar-only),
1,729--2,084 ns/op (array-heavy), 3,108--3,246 ns/op (mixed-domain), and
125--161 ns/op (recovered), with 800/9, 720/11, 832/15, and 400/3 B/op and
allocs/op, respectively.

The exact command for this microbenchmark was:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./internal/analyze -run '^$' -bench '^BenchmarkProcedureApplicabilityPlanning$' -benchmem -benchtime=100x -count=3
```

The `-count=1` CPU/allocation profile completed in 23.83 s with 84.88 s of CPU
samples. Its flat CPU top ten were `runtime.semasleep` (13.57%),
`runtime.scanObject` (4.92%), `runtime.tryDeferToSpanScan` (4.50%),
`regexp.(*Regexp).tryBacktrack` (3.19%), `aeshashbody` (2.95%),
`runtime.spanClass.sizeclass` (2.58%), `runtime.findObject` (2.21%),
`internal/runtime/maps.(*Iter).Next` (2.14%), `strings.ToLower` (1.89%), and
`runtime.scanObjectsSmall` (1.87%).

The alloc-space top ten were `cloneHTTPState` (5,755.43 MB),
`cloneArrayState` (5,420.14 MB), `internal/bytealg.MakeNoZero` (2,049.05 MB),
`meetArrayState` (1,694.06 MB), `arrayVariables.func1` (1,106.16 MB),
`strings.Fields` (1,098.09 MB), `cfg.buildQueryIndex` (824.41 MB),
`constexpr.NewValues` (520.52 MB), `cfg.Graph.Dominators` (426.53 MB), and
`dcFlowState.clone` (425.54 MB). The alloc-object top ten were
`internal/bytealg.MakeNoZero` (58,031,404), `strings.Fields` (18,443,983),
`reflect.unsafe_New` (17,694,988), `cloneArrayState` (4,887,930),
`strings.genSplit` (4,482,038), tree-sitter `GoString` (3,883,006),
`cloneApplicationStateSnapshot` (3,746,639), `cfg.buildQueryIndex` (3,598,583),
`cloneHTTPState` (3,361,323), and `meetArrayState` (2,406,706).

The unprofiled, bounded `independent/2000` scheduling benchmark changed from a
pre-instrumentation median of 1.391 s to 1.382 s after instrumentation
(-0.6%). Median allocation changed from 1,469,546,792 B/op and 11,452,572
allocs/op to 1,470,062,104 B/op and 11,450,001 allocs/op. This is within the
local 2% disabled-path threshold; timing remains observational rather than a
CI assertion.

Use the same five-sample benchmark convention for comparisons where practical;
the profile-producing command itself remains a single `-count=1` run so profile
files are not overwritten. These are developer observations for same-machine
comparison, not fixed CI thresholds.

### Analyzer procedure-view allocation record (#698)

Issue #698 removes the complete declaration, statement, expression, call, and
access slice projections from the normal analyzer procedure path. The
projection-only benchmark keeps IR fixture construction outside the timed
region and compares view materialization with iteration over the cached view.
The local run used Go 1.26.6 on the same Windows 11 / Intel Core i7-12700
developer machine as the preceding corpus records.

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./internal/analyze -run '^$' -bench '^BenchmarkParsedFileProcedureProjection$' -benchmem -benchtime=1x -count=3
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./internal/analyze -run '^$' -bench 'BenchmarkSingleModuleSynthetic/(independent|declarations|chain|cfg-independent)/2000-' -benchmem -benchtime=1x -count=1
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./internal/staticanalysis/corpus -run '^$' -bench '^BenchmarkRealWorldCorpus/ronecone/analyze-only$' -benchmem -benchtime=1x -count=5
```

The projection-only median observations were:

| procedures | materialize `B/op` / `allocs/op` | cached iteration `B/op` / `allocs/op` |
| ---------: | -------------------------------: | ------------------------------------: |
|        100 |                  175,360 / 1,501 |                            41,440 / 2 |
|        500 |                  860,416 / 7,501 |                           188,896 / 2 |
|      1,000 |               1,720,832 / 15,001 |                           377,312 / 2 |
|      2,000 |               3,441,664 / 30,001 |                           754,144 / 2 |

The previous 500-procedure projection baseline was 1,461,257 B/op and 10,002
allocs/op; the new normal-path materialization is 41% lower in bytes and 25%
lower in allocation count. End-to-end synthetic 2,000-procedure observations
were 1,481,884,216 B/op and 11,500,704 allocs/op for `independent`,
1,510,002,272 / 11,989,668 for `chain`, 3,579,487,696 / 1,045,050 for
`declarations`, and 7,978,821,776 / 8,008,097 for `cfg-independent`.
Findings and fact-build counters remained unchanged for these fixtures.

The five ROneCOne `analyze-only` samples produced 23.135, 24.952, 25.329,
33.490, and 37.966 seconds, with 26,642,309,208--26,661,017,936 B/op and
177,339,816--177,458,969 allocs/op. The median was 25,328,943,600 ns/op,
26,649,626,264 B/op, and 177,360,666 allocs/op. The wide timing spread is a
developer-machine observation; allocation volume did not regress. All five
samples reported 1,565 procedures and 1,565 procedure fact builds, with the
same 510 findings and planner counters.

The single profile run used the documented `-cpuprofile`/`-memprofile` command.
Its allocation-space top entries were state/CFG/data-flow working storage
(`cloneHTTPState`, `cloneArrayState`, `MakeNoZero`, and `meetArrayState`), not
procedure projection copies. Allocation-object top entries were likewise
dominated by `MakeNoZero`, reflection, `strings.Fields`, state clones, and CFG
query construction. The in-use profile was startup/runtime dominated, so these
profiles are evidence that the removed projection copies are no longer a
material corpus allocation leaf rather than a claim that all corpus allocation
has been eliminated.

## Verification requirements

Manifest tests cover valid v2 data and rejection of unknown fields, unsupported
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
materialization of all 16 vendored projects, generated profile configs,
document-module placement, source-map remapping, temporary-path exclusion,
disabled-project skips, partial-result preservation, and cleanup after both
success and failure. A representative third-party project is analyzed through
the production lint/analyze implementations. Snapshot tests additionally cover
the exact JSONL identity and field order, project-relative path remapping,
deterministic ordering across repeated runs, duplicate-row multiplicity,
prose exclusion, empty snapshots, added/removed diagnostic failures, missing
snapshot failures, verify-only behavior, explicit update generation, and the
no-partial-update guarantee when any project or surface fails.

Corpus synchronization and execution do not change static-analysis semantics,
public CLI/API output, Excel/VBE oracle behavior, or Go dependency-inventory
checks. Those suites remain separate from synchronization and may run fully
offline.

Issue #548 added focused completion regressions at the analyzer boundary. Type
inference tracks active `Set` assignments so a self-referential RHS terminates
while still consulting an earlier same-scope assignment. Project effect
propagation uses indexed evidence and uncertainty membership while preserving
first-seen ordering. Procedure dataflow uses a deterministic reverse-postorder
priority worklist built from pre-indexed CFG adjacency, with iterative traversal
to avoid analyzer-stack growth. The dataflow regression analyzes the vendored
`LibJSON.ParseChars` procedure and also covers a synthetic 8,000-block-deep,
4,000-block-wide CFG. These are source-independent analyzer rules; no corpus
project name, file omission, diagnostic suppression, or relaxed timeout is
used to make a project complete.

The reverse-postorder change preserves the source, sink, sanitizer, join, and
path-witness contracts. Snapshot review removed three pre-existing `VBA224`
rows that arose only from transient, partially propagated FIFO states; all
remaining diagnostics are stable across repeated verify-only runs.

## Related

- `docs/adr/ADR-0029-vendored-static-analysis-corpus.md`
- `docs/adr/ADR-0030-third-party-corpus-workspaces.md`
- `docs/adr/ADR-0031-deterministic-static-analysis-snapshots.md`
- `docs/adr/ADR-0032-reviewed-static-analysis-corpus-evidence.md`
- `docs/adr/ADR-0024-shared-static-analysis-rule-registry.md`
- `docs/adr/ADR-0026-local-vbe-oracle-evidence.md`
- Issue #530
- Issue #531
- Issue #537
- Issue #548
