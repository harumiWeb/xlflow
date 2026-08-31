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

## Validation and Adapters

The model is a passive value contract and performs no construction-time
validation. The analysis entry point that consumes it is responsible for
reporting unsupported kinds, duplicate identities, or other invalid inputs.

Filesystem discovery and source loading remain adapter responsibilities. The
existing `symbols.SourceFile` is a discovery descriptor containing a path and
inferred kind; it is not a loaded source project. Issue #641 will convert such
descriptors into this model after reading source bytes. Issue #642 will add the
common analysis entry point.

This contract does not introduce a virtual filesystem, change CLI or LSP wire
formats, refactor the analyzer, or provide browser/Wasm integration.
