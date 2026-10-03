package procedureir

import "strings"

// BuildProjectResolver constructs the canonical resolver for a project IR
// snapshot. External symbols are appended as provided, and completeness is
// passed through so incomplete project snapshots keep their fail-open behavior.
func BuildProjectResolver(documents []DocumentIR, externalSymbols []ResolverSymbol, complete bool) Resolver {
	symbolCount := len(externalSymbols)
	for _, document := range documents {
		symbolCount += len(document.Declarations) + len(document.Procedures)
	}
	resolverSymbols := make([]ResolverSymbol, 0, symbolCount)
	for _, document := range documents {
		module := strings.TrimSpace(document.ModuleName)
		for _, declaration := range document.Declarations {
			resolverSymbols = append(resolverSymbols, ResolverSymbol{
				Name: declaration.Name, Type: declaration.Type, Module: module, ModuleKind: document.ModuleKind,
				Kind: declaration.Kind, Visibility: declaration.Visibility, File: document.Path,
				Line: declaration.Range.StartLine, Parent: declaration.Parent, Recovered: declaration.Recovered,
				IsArray: declaration.IsArray, IsConst: declaration.IsConst,
				ValueShape:          declaration.ValueShape,
				ConditionalBranches: append([]ConditionalBranch(nil), declaration.ConditionalBranches...),
			})
		}
		for _, procedure := range document.Procedures {
			resolverSymbols = append(resolverSymbols, ResolverSymbol{
				Name: procedure.Symbol.Name, Type: procedure.Symbol.ReturnType, Module: module, ModuleKind: document.ModuleKind,
				Kind: string(procedure.Symbol.Kind), Visibility: procedure.Symbol.Visibility, File: document.Path,
				Line: procedure.Symbol.DeclarationRange.StartLine, Recovered: procedure.Symbol.Recovered,
				IsArray: procedure.Symbol.IsArray, ValueShape: procedure.Symbol.ValueShape,
				ConditionalBranches: append([]ConditionalBranch(nil), procedure.Symbol.ConditionalBranches...),
			})
		}
	}
	resolverSymbols = append(resolverSymbols, externalSymbols...)
	return NewResolverWithCompleteness(resolverSymbols, complete)
}
