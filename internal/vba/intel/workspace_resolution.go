package intel

import (
	"context"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
)

// WorkspaceResolutionView indexes an immutable request-scoped symbol snapshot.
// Query results are copied so callers retain the historical ownership contract.
type WorkspaceResolutionView struct {
	all       []Symbol
	exact     map[string][]int
	qualified map[string][]int
	module    map[string][]int
	kind      map[string][]int
	exactKeys []string
}

func NewWorkspaceResolutionView(symbols []Symbol) *WorkspaceResolutionView {
	view := &WorkspaceResolutionView{
		all:       append([]Symbol(nil), symbols...),
		exact:     make(map[string][]int),
		qualified: make(map[string][]int),
		module:    make(map[string][]int),
		kind:      make(map[string][]int),
	}
	for i, symbol := range view.all {
		view.exact[normalizeWorkspaceResolutionKey(symbol.Name)] = append(view.exact[normalizeWorkspaceResolutionKey(symbol.Name)], i)
		qualified := qualifiedName(symbol.Module, symbol.Name)
		view.qualified[normalizeWorkspaceResolutionKey(qualified)] = append(view.qualified[normalizeWorkspaceResolutionKey(qualified)], i)
		view.module[normalizeWorkspaceResolutionKey(symbol.Module)] = append(view.module[normalizeWorkspaceResolutionKey(symbol.Module)], i)
		view.kind[normalizeWorkspaceResolutionKey(symbol.Kind)] = append(view.kind[normalizeWorkspaceResolutionKey(symbol.Kind)], i)
	}
	view.exactKeys = make([]string, 0, len(view.exact))
	for key := range view.exact {
		view.exactKeys = append(view.exactKeys, key)
	}
	sort.Strings(view.exactKeys)
	return view
}

func (v *WorkspaceResolutionView) Query(query WorkspaceSymbolQuery) []Symbol {
	if v == nil {
		return nil
	}
	key := normalizeWorkspaceResolutionKey(query.Text)
	var indexes []int
	switch query.Mode {
	case WorkspaceSymbolQueryExact:
		indexes = v.exact[key]
	case WorkspaceSymbolQueryQualified:
		indexes = v.qualified[key]
	case WorkspaceSymbolQueryModule:
		indexes = v.module[key]
	case WorkspaceSymbolQueryKind:
		indexes = v.kind[key]
	case WorkspaceSymbolQueryPrefix:
		if key == "" {
			return append([]Symbol(nil), v.all...)
		}
		start := sort.SearchStrings(v.exactKeys, key)
		for i := start; i < len(v.exactKeys) && strings.HasPrefix(v.exactKeys[i], key); i++ {
			indexes = append(indexes, v.exact[v.exactKeys[i]]...)
		}
	default:
		if key == "" {
			return append([]Symbol(nil), v.all...)
		}
		for i, symbol := range v.all {
			if strings.Contains(strings.ToLower(symbol.Name), key) || strings.Contains(strings.ToLower(qualifiedName(symbol.Module, symbol.Name)), key) {
				indexes = append(indexes, i)
			}
		}
	}
	result := make([]Symbol, 0, len(indexes))
	for _, index := range indexes {
		result = append(result, v.all[index])
	}
	sort.SliceStable(result, func(i, j int) bool {
		if result[i].File != result[j].File {
			return result[i].File < result[j].File
		}
		if result[i].Range.Start.Line != result[j].Range.Start.Line {
			return result[i].Range.Start.Line < result[j].Range.Start.Line
		}
		if result[i].Range.Start.Character != result[j].Range.Start.Character {
			return result[i].Range.Start.Character < result[j].Range.Start.Character
		}
		return result[i].Name < result[j].Name
	})
	return result
}

func (a Analyzer) withRequestWorkspaceResolution(ctx context.Context, open []Document) Analyzer {
	if a.WorkspaceSymbolsSnapshotFunc == nil || ctx.Err() != nil {
		return a
	}
	symbols, err := a.WorkspaceSymbolsSnapshotFunc(open)
	if err != nil || ctx.Err() != nil {
		return a
	}
	a.visibleDeclarations = visibleAssignmentDeclarations(symbols)
	view := NewWorkspaceResolutionView(symbols)
	if recorder := analysisstats.FromContext(ctx); recorder != nil {
		recorder.Add("workspace_resolution_views", 1)
	}
	a.WorkspaceSymbolQueryFunc = func(_ []Document, query WorkspaceSymbolQuery) ([]Symbol, error) {
		return view.Query(query), nil
	}
	a.WorkspaceSymbolsFunc = nil
	return a
}

func visibleAssignmentDeclarations(symbols []Symbol) map[string]bool {
	visible := make(map[string]bool)
	for _, symbol := range symbols {
		if symbol.Parent != "" || !strings.EqualFold(symbol.Visibility, "Public") || !strings.EqualFold(symbol.ModuleKind, "standard") {
			continue
		}
		switch strings.ToLower(symbol.Kind) {
		case "module_variable", "const", "type", "enum":
			visible[symbol.Name] = true
		}
	}
	return visible
}

func normalizeWorkspaceResolutionKey(value string) string {
	return strings.ToLower(strings.TrimSpace(value))
}
