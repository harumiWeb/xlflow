# In-Memory VBA Source Project

This specification defines the protocol-neutral source input model introduced
by issue #640. The model is the common representation that filesystem, editor,
test, and future embedded adapters can construct before invoking static
analysis. ADR-0050 records the architectural rationale.

## Model

`internal/vba/sourceproject` exposes these Go-owned values:

```go
type ModuleKind string

const (
    ModuleKindStandard ModuleKind = "standard"
    ModuleKindClass    ModuleKind = "class"
    ModuleKindForm     ModuleKind = "form"
    ModuleKindDocument ModuleKind = "document"
)

type SourceFile struct {
    Path       string
    Source     []byte
    ModuleKind ModuleKind
    IsTest     bool
}

type SourceProject struct {
    Files []SourceFile
}
```

The module kinds retain the existing analyzer vocabulary. Workbook and
worksheet code modules both use `document`; their logical path and VBA source
retain their individual module identity. UserForm code uses `form` regardless
of whether a filesystem adapter acquired it from an exported `.frm` file or a
sidecar source file.

A test module remains a VBA `standard` module and sets `IsTest`. Test role is
orthogonal to VBA component semantics so project name resolution continues to
treat test helpers and procedures as standard-module declarations.

## Path and Source Contract

`Path` is a caller-provided logical identity used for diagnostics and as a
module-name fallback. It can be relative, absolute, or virtual and does not
need to name an existing file. The model does not clean, resolve, stat, read,
or classify the path.

`Source` contains the exact source bytes supplied by the caller. The model does
not read source from `Path` and does not make an implicit copy. Callers must not
mutate the bytes while a consumer is using the project. Consumers that retain
source beyond a call boundary own any snapshotting they require.

The model package does not import configuration, parser, LSP, Excel, COM, or
OS-specific packages.

## Filesystem Adapter

`internal/vba/sourceprojectfs` is the filesystem-backed adapter for batch
static analysis. It discovers configured module, class, form, and workbook
roots through the canonical `symbols` discovery contract, also includes the
legacy top-level `tests` tree, applies an optional path filter, reads the
selected files, and returns a `SourceProject`.

The adapter preserves source bytes exactly and uses the module kind assigned by
canonical discovery. A standard module discovered through the legacy `tests`
tree sets `IsTest`; class, form, and document semantics remain represented by
their module kind. When a path is reachable through both production discovery
and `tests`, the production entry takes precedence. Results are deduplicated
and sorted deterministically by path before source is read.

Filesystem reads and deduplication use absolute physical paths internally.
Returned logical paths and path-filter inputs preserve the form implied by
`RootDir`: absolute roots produce absolute paths, while relative or empty roots
produce relative paths as in the existing batch analyzer contract.

UserForm source selection follows the configured code-source mode. In sidecar
mode an authoritative `forms/code/<Name>.bas` entry is classified as `form` and
the matching exported `.frm` code is not loaded. A `.frm` remains eligible when
there is no authoritative sidecar.

The optional path filter runs before `os.ReadFile`, so excluded files are not
loaded into the returned project. Missing configured roots and a missing
legacy `tests` tree contribute no files; discovery, read, and cancellation
errors are returned to the caller.

## Source Encoding

Tracked `.bas`, `.cls`, and `.frm` files use UTF-8 without a BOM. The shared
`internal/vba/sourceencoding` package owns BOM detection, UTF-8 validation,
diagnostic byte positions, managed-root discovery, and strict CP932 decoding
for the explicit conversion command. It includes `.frm` files even when
`[userform].code_source = "sidecar"` and excludes binary `.frx` companions.

The parser's normal, context/realtime, incremental, and file-loading entry
points validate source bytes before passing them to tree-sitter. Formatter and
other parser consumers use the same checked entry point. Invalid bytes return
a typed source-encoding error containing the logical path and first offending
position; callers must not replace invalid bytes or silently convert them.
Filesystem adapters may preserve the original bytes in `Source`, but any
parser/analyzer operation that consumes them fails with that typed error.

## Validation and Adapters

The model is a passive value contract and performs no construction-time
validation. The analysis entry point that consumes it is responsible for
reporting unsupported kinds, duplicate identities, or other invalid inputs.

`internal/analyze.Analyzer.AnalyzeProject` is the common batch-analysis entry
point:

```go
func (a Analyzer) AnalyzeProject(
    ctx context.Context,
    project sourceproject.SourceProject,
) (Result, error)
```

It treats the supplied files as the complete project and never performs source
discovery or reads `SourceFile.Path`. The input is validated before parsing:
module kinds must be one of the declared constants, paths must be non-empty,
normalized logical identities must be unique, and `IsTest` is valid only for
standard modules. The caller-owned file order and source bytes are not mutated;
the analyzer uses a deterministic internal ordering while retaining each
caller-supplied logical path for diagnostics and errors. An empty project is a
successful analysis with zero analyzed files.

`Analyzer.RunResultContext` remains the filesystem adapter. It loads a
`SourceProject` through `sourceprojectfs.LoadContext` and delegates parsing,
IR/CFG construction, project resolution/effects, diagnostics, and finalization
to the same analysis core. Its existing `PathFilter` behavior, including the
filesystem-only supplemental symbol lookup needed for excluded project
candidates, remains adapter behavior and is not applied by `AnalyzeProject`.

Inline suppressions in the common core are parsed from each supplied
`SourceFile.Source`, so a virtual path does not trigger a second filesystem
read. The broader classification of diagnostics that need external data (for
example UserForm Designer/FRX metadata) remains the responsibility of issue
#644.

Filesystem discovery and source loading remain adapter responsibilities. The
existing `symbols.SourceFile` is a discovery descriptor containing a path and
inferred kind; it is not a loaded source project. The filesystem adapter
converts those descriptors into this model after reading source bytes. Issue
#642 adds the common analysis entry point; issue #644 owns the remaining
filesystem-free diagnostic and suppression capability policy.

This contract does not introduce a virtual filesystem, change CLI or LSP wire
formats, or provide browser/Wasm integration.
