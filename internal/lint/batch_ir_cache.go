package lint

import (
	"crypto/sha256"
	"path/filepath"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// batchIRCacheEntry holds invocation-local owned Go values, never AST handles.
type batchIRCacheEntry struct {
	sourceSHA256 [32]byte
	ir           procedureir.DocumentIR
}

func (l Linter) rememberBatchIR(path string, source []byte, ir procedureir.DocumentIR) {
	if l.batchIRCache == nil || ir.Parse.HasError || ir.Parse.HasMissing {
		return
	}
	l.batchIRCache[filepath.Clean(path)] = batchIRCacheEntry{
		sourceSHA256: sha256.Sum256(source), ir: procedureir.CloneDocumentIR(ir),
	}
}

func (l Linter) reuseBatchIR(path string, source []byte, module, kind string) (procedureir.DocumentIR, bool) {
	entry, ok := l.batchIRCache[filepath.Clean(path)]
	if !ok || module == "" || kind == "" || entry.sourceSHA256 != sha256.Sum256(source) ||
		entry.ir.ModuleName != module || entry.ir.ModuleKind != kind {
		return procedureir.DocumentIR{}, false
	}
	// RootDir only controls the builder's display paths. Exact module identity
	// was checked above; never reinterpret event classification after building.
	display := path
	if strings.TrimSpace(l.RootDir) != "" && filepath.IsAbs(path) {
		if relative, err := filepath.Rel(l.RootDir, path); err == nil {
			display = relative
		}
	}
	display = filepath.ToSlash(display)
	ir := procedureir.CloneDocumentIR(entry.ir)
	ir.Path = display
	for i := range ir.Procedures {
		for j := range ir.Procedures[i].Calls {
			ir.Procedures[i].Calls[j].File = display
		}
	}
	for i := range ir.TypeReferences {
		ir.TypeReferences[i].File = display
	}
	return ir, true
}
