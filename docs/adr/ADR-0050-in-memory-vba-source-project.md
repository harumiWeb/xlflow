# ADR-0050: Protocol-Neutral In-Memory VBA Source Project

## Status

`accepted`

## Background

The batch analyzer discovers configured source directories, reads VBA files,
classifies their component kinds, and performs static analysis in one entry
point. LSP buffers and several lower analysis layers already accept source in
memory, but there is no common project-level input that callers can construct
without a filesystem.

Issue #599 requires one production analysis boundary for the CLI, LSP, tests,
and future browser or embedded consumers. Issue #640 establishes only the
source model needed by that work; filesystem loading and analyzer extraction
remain separate follow-up issues.

Existing models have narrower ownership. `symbols.SourceFile` describes a file
found by configured filesystem discovery and has no source bytes.
`intel.Document` carries LSP-oriented URI, version, and snapshot state. The
packer's `vbaproject.Project` models binary VBA project streams and references
for read-modify-write operations.

## Decision

Add `internal/vba/sourceproject` as a dependency-light package containing a
`SourceProject` collection and caller-supplied `SourceFile` values. Each file
retains a logical path, explicit source bytes, and one of the existing
`standard`, `class`, `form`, or `document` component kinds.

Treat test role independently from VBA component kind. A test module remains a
standard module and carries `IsTest`, preserving the standard-module lookup
semantics used by test discovery and project resolution.

Treat paths as logical source identities rather than proof of filesystem
existence. The model performs no discovery, reads, path normalization, module
classification, parsing, or validation. It contains no configuration, LSP,
Excel, COM, or OS-specific dependency. Callers own source byte slices and keep
them immutable while a consumer uses the project.

Future filesystem and editor adapters construct this model. The future common
analysis entry point validates and consumes it through the existing parser,
procedure IR, CFG, and semantic implementations.

Type-aware analysis receives its type metadata through an explicit capability
that carries both the caller-owned `vbadb.DB` and whether the view is complete.
`AnalyzeProject` uses only the embedded built-in database when no capability is
provided and marks that view incomplete; the filesystem-backed adapter may
instead load and overlay the generated TypeLib database. The common semantic
implementation is shared in both cases, while incomplete views fail open for
diagnostics that would otherwise infer external type absence.

The filesystem-backed adapter treats generated metadata as complete only when a
caller that has build metadata supplies the expected TypeLib generator version
and the manifest matches it. A mismatch does not discard generated members:
positive type resolution remains available, but completeness-dependent absence
diagnostics fail open and the existing TypeDB load warning is retained. Callers
without build metadata keep the version-agnostic loader contract. Standalone
realtime analysis exposes the same opt-in expected-version path; explicit
`TypeDatabase` injection remains authoritative and bypasses runtime loading.

Keep `AnalyzeProject` source-only after the common boundary is introduced.
Source-local data comes only from the caller-supplied `SourceFile` values, and
virtual-project metadata may be supplied by another logical file in the same
`SourceProject`. The analyzer must not open a logical path or discover a
sidecar artifact. In particular, a supplied `.frm` can provide UserForm
control metadata for its code-behind module, while an unprovided or
`.frx`-dependent Designer view is incomplete. `VBA220` then remains
conservative and `Result.Warnings` reports the structured
`analysis_capability_unavailable` / `userform_control_metadata` capability for
the affected logical file. The warning is emitted in deterministic source
order. Filesystem-backed `RunResultContext` keeps its existing sidecar and
FRX behavior; LSP and realtime adapter contracts are outside this decision.

Workbook metadata follows the same explicit-capability boundary. A caller may
provide an immutable worksheet visible-name to CodeName catalog to the
analyzer. The filesystem-backed adapter may populate that catalog from the
configured saved OOXML workbook, but `AnalyzeProject` never opens a workbook
or interprets a logical source path to obtain it. A missing, malformed, or
ambiguous catalog fails open for `VBA260` and is reported as the structured
`analysis_capability_unavailable` / `worksheet_codename_catalog` warning when
the rule is enabled. TypeLib metadata remains separate because worksheet names
and CodeNames are workbook-instance facts, not Excel object-model type facts.

## Consequences

Positive consequences:

- Multi-module VBA projects can be represented completely in memory.
- Source acquisition and protocol lifecycle stay outside the analysis input.
- Explicit component kinds avoid filesystem inspection and extension guesses.
- Test modules retain both their test role and standard-module semantics.
- `AnalyzeProject` cannot accidentally make a virtual path observable through
  host-side Designer, FRX, or workspace-symbol reads.
- External UserForm metadata has an explicit conservative diagnostic and
  structured warning contract instead of a hidden filesystem fallback.
- Workbook-instance worksheet identity is injected explicitly and cannot make
  virtual source paths observable through an implicit workbook read.

Negative consequences:

- The source model temporarily coexists with discovery and LSP document types
  until the adapters in issues #641 and #642 adopt it.
- Callers must preserve source byte immutability or make their own snapshot.
- Validation cannot occur until a consumer applies its input requirements.
- Callers that need complete UserForm control metadata must supply the relevant
  `.frm` source (and any binary metadata through a future explicit capability);
  `AnalyzeProject` does not infer it from `RootDir`.

## Alternatives Considered

1. **Extend `symbols.SourceFile` with source bytes.** Rejected because its
   package and meaning are tied to configured filesystem discovery.
2. **Reuse `intel.Document`.** Rejected because URI, revision, and analysis
   snapshot lifecycle are editor concerns rather than source-project inputs.
3. **Reuse the packer's binary project model.** Rejected because binary stream,
   reference, and protection metadata are unrelated to static source analysis.
4. **Represent tests as a fifth module kind.** Rejected because xlflow tests are
   standard VBA modules and must retain standard-module resolution behavior.

## Evidence

- Requirements: xlflow issues #599 and #640.
- Existing source discovery: `internal/vba/symbols/symbols.go`.
- Existing LSP source ownership: `internal/vba/intel/intel.go` and
  `internal/lspserver/workspace_analysis_index.go`.
- Existing analysis input construction: `internal/analyze/analyzer.go`.
- Model contract and tests: `docs/specs/vba-source-project.md` and
  `internal/vba/sourceproject`.
- Generator-version compatibility follow-up: issue #813,
  `internal/typedb/typedb.go`, `internal/analyze/analyzer.go`, and
  `internal/analyze/analyzer_test.go`.
- Filesystem-free diagnostic capability policy: issue #644,
  `internal/analyze/event_reentry.go`, `internal/vba/intel/intel.go`, and
  `docs/specs/vba-source-project.md`.
- Excel semantic inspections and worksheet identity capability: issue #827,
  `internal/ooxml/workbook.go`, `internal/analyze/worksheet_codename.go`, and
  `docs/specs/excel-semantic-inspections.md`.

## Supersedes

- None.

## Superseded by

- None.

## Related

- `docs/adr/ADR-0014-reusable-vba-lsp-server.md`
- `docs/adr/ADR-0021-procedure-analysis-ir.md`
- xlflow issues #641, #642, #643, #644, #645, #813, and #827
