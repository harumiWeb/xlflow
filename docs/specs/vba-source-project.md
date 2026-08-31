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

## Validation and Adapters

The model is a passive value contract and performs no construction-time
validation. The analysis entry point that consumes it is responsible for
reporting unsupported kinds, duplicate identities, or other invalid inputs.

Filesystem discovery and source loading remain adapter responsibilities. The
existing `symbols.SourceFile` is a discovery descriptor containing a path and
inferred kind; it is not a loaded source project. The filesystem adapter
converts those descriptors into this model after reading source bytes. Issue #642
will add the common analysis entry point, while issue #644 owns
filesystem-free diagnostic and suppression behavior.

This contract does not introduce a virtual filesystem, change CLI or LSP wire
formats, refactor the analyzer, or provide browser/Wasm integration.
