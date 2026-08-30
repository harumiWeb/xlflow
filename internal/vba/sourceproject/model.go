// Package sourceproject defines the filesystem-independent input model for a
// VBA source project.
package sourceproject

// ModuleKind identifies the VBA component semantics that cannot be inferred
// reliably from source text alone.
type ModuleKind string

const (
	ModuleKindStandard ModuleKind = "standard"
	ModuleKindClass    ModuleKind = "class"
	ModuleKindForm     ModuleKind = "form"
	ModuleKindDocument ModuleKind = "document"
)

// SourceFile is one caller-supplied VBA source module.
//
// Path is a logical identity used for diagnostics and module-name fallback. It
// does not need to exist on disk. Source is caller-owned and must not be
// mutated while a consumer is using the project. IsTest describes the role of
// a standard module without changing its VBA module semantics.
type SourceFile struct {
	Path       string
	Source     []byte
	ModuleKind ModuleKind
	IsTest     bool
}

// SourceProject is a caller-supplied collection of VBA source modules.
type SourceProject struct {
	Files []SourceFile
}
