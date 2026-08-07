package intel

import (
	"context"
	"fmt"
	"strings"

	staticrules "github.com/harumiWeb/xlflow/internal/staticanalysis/rules"
	"github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// LocalTypeNameDiagnosticsContext reports local As-type names that the
// production VBA type/project resolver cannot resolve. The resolver is
// deliberately fail-closed: an incomplete workspace lookup is not treated as
// proof that a name is missing.
func (a Analyzer) LocalTypeNameDiagnosticsContext(ctx context.Context, doc Document) []Diagnostic {
	if ctx.Err() != nil || a.DB == nil || a.TypeDBResolutionIncomplete {
		return nil
	}
	parsed, closeParsed, err := parsedDocumentForDocument(doc)
	if err != nil {
		return nil
	}
	defer closeParsed()
	ir, err := procedureIRForDocumentContext(ctx, doc, a.RootDir, parsed)
	if err != nil || ctx.Err() != nil {
		return nil
	}
	return a.localTypeNameDiagnostics(ctx, doc, ir)
}

func (a Analyzer) localTypeNameDiagnostics(ctx context.Context, doc Document, ir procedureir.DocumentIR) []Diagnostic {
	if a.DB == nil || a.TypeDBResolutionIncomplete || ctx.Err() != nil {
		return nil
	}
	var out []Diagnostic
	for _, procedure := range ir.Procedures {
		if ctx.Err() != nil {
			return nil
		}
		for _, reference := range ir.TypeReferences {
			if reference.Kind != "uses_type" || reference.Caller == nil ||
				!sameProcedureReference(*reference.Caller, procedure) {
				continue
			}
			for _, declaration := range procedure.Declarations {
				if declaration.Scope != procedureir.ScopeLocal || declaration.Kind != "variable" ||
					!astRangeContains(declaration.Range, reference.Range) {
					continue
				}
				resolved, complete := a.resolveLocalTypeName(doc, reference.Target)
				if !complete || resolved {
					break
				}
				severity := "error"
				if metadata, ok := staticrules.Lookup("VBA229"); ok {
					severity = string(metadata.DefaultSeverity)
				}
				out = append(out, Diagnostic{
					Code:     "VBA229",
					Severity: severity,
					Source:   "xlflow",
					Message:  fmt.Sprintf("Unresolved local As type name %s.", strings.TrimSpace(reference.Target)),
					Range: Range{
						Start: positionForDocumentByteOffset(doc, reference.Range.StartByte),
						End:   positionForDocumentByteOffset(doc, reference.Range.EndByte),
					},
					Rule:       "VBA229",
					Confidence: "high",
				})
				break
			}
		}
	}
	return out
}

func sameProcedureReference(reference procedureir.ProcedureRef, procedure procedureir.ProcedureIR) bool {
	if reference.QualifiedName != "" && procedure.Symbol.QualifiedName != "" {
		return strings.EqualFold(reference.QualifiedName, procedure.Symbol.QualifiedName)
	}
	return strings.EqualFold(reference.Name, procedure.Symbol.Name) &&
		strings.EqualFold(string(reference.Kind), string(procedure.Symbol.Kind))
}

func astRangeContains(outer, inner ast.Range) bool {
	return outer.StartByte <= inner.StartByte && outer.EndByte >= inner.EndByte
}

func (a Analyzer) resolveLocalTypeName(doc Document, target string) (resolved, complete bool) {
	target = strings.TrimSpace(target)
	if target == "" {
		return false, true
	}
	if isBuiltinLocalTypeName(target) {
		return true, true
	}
	if _, ok := a.DB.ResolveType(target); ok {
		return true, true
	}
	mode := WorkspaceSymbolQueryExact
	if strings.Contains(target, ".") {
		mode = WorkspaceSymbolQueryQualified
	}
	symbols, err := a.WorkspaceSymbolsQuery([]Document{doc}, WorkspaceSymbolQuery{Text: target, Mode: mode})
	if err != nil {
		return false, false
	}
	if mode == WorkspaceSymbolQueryQualified {
		for _, symbol := range symbols {
			if projectTypeSymbol(symbol) {
				return true, true
			}
		}
		return false, true
	} else if short := shortTypeName(target); short != target {
		// The production workspace view indexes qualified names separately, but
		// a custom provider may return only the short-name candidates.
		for _, symbol := range symbols {
			if strings.EqualFold(symbol.Name, short) && projectTypeSymbol(symbol) {
				return true, true
			}
		}
		return false, true
	}
	for _, symbol := range symbols {
		if projectTypeSymbol(symbol) {
			return true, true
		}
	}
	return false, true
}

func projectTypeSymbol(symbol Symbol) bool {
	switch strings.ToLower(strings.TrimSpace(symbol.Kind)) {
	case "type", "enum", "class":
		return true
	case "module":
		return strings.EqualFold(symbol.ModuleKind, "class")
	default:
		return false
	}
}

func isBuiltinLocalTypeName(name string) bool {
	name = strings.ToLower(strings.TrimSpace(name))
	for _, builtin := range builtinVBATypeNames {
		if name == strings.ToLower(builtin) {
			return true
		}
	}
	return false
}
