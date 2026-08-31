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

## Consequences

Positive consequences:

- Multi-module VBA projects can be represented completely in memory.
- Source acquisition and protocol lifecycle stay outside the analysis input.
- Explicit component kinds avoid filesystem inspection and extension guesses.
- Test modules retain both their test role and standard-module semantics.

Negative consequences:

- The source model temporarily coexists with discovery and LSP document types
  until the adapters in issues #641 and #642 adopt it.
- Callers must preserve source byte immutability or make their own snapshot.
- Validation cannot occur until a consumer applies its input requirements.

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

## Supersedes

- None.

## Superseded by

- None.

## Related

- `docs/adr/ADR-0014-reusable-vba-lsp-server.md`
- `docs/adr/ADR-0021-procedure-analysis-ir.md`
- xlflow issues #641, #642, #644, and #645
