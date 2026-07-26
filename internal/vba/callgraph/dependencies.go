package callgraph

import (
	"fmt"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/calls"
)

// DependencyNode is one stable entity in the project dependency projections.
// Kind is procedure, module, or type.
type DependencyNode struct {
	ID         string   `json:"id"`
	Kind       string   `json:"kind"`
	Name       string   `json:"name"`
	Module     string   `json:"module,omitempty"`
	ModuleKind string   `json:"module_kind,omitempty"`
	File       string   `json:"file,omitempty"`
	Location   Location `json:"location"`
}

// DependencyEvidence preserves the source fact that produced a dependency.
// Source and Target identify the underlying procedure pair for an aggregated
// module call edge when applicable.
type DependencyEvidence struct {
	Kind     string   `json:"kind"`
	File     string   `json:"file"`
	Location Location `json:"location"`
	Source   string   `json:"source,omitempty"`
	Target   string   `json:"target,omitempty"`
}

// DependencyEdge is a confirmed, typed, project-local dependency.
type DependencyEdge struct {
	Kind     string               `json:"kind"`
	From     string               `json:"from"`
	To       string               `json:"to"`
	Evidence []DependencyEvidence `json:"evidence"`
}

// UncertainEdge is intentionally separate from confirmed edges so consumers
// cannot mistake dynamic or unresolved calls for project-local dependencies.
type UncertainEdge struct {
	Status     string             `json:"status"`
	From       string             `json:"from"`
	Callee     string             `json:"callee"`
	Candidates []calls.Candidate  `json:"candidates,omitempty"`
	Evidence   DependencyEvidence `json:"evidence"`
}

// DependencyRequest filters the outward dependency view to one exact module.
// An empty Module returns the whole project graph.
type DependencyRequest struct {
	Module string
}

// DependencyResult contains direct procedure and aggregated module projections
// together with conservative type dependencies and raw uncertainty.
type DependencyResult struct {
	Target         string           `json:"target"`
	Nodes          []DependencyNode `json:"nodes"`
	Edges          []DependencyEdge `json:"edges"`
	UncertainEdges []UncertainEdge  `json:"uncertain_edges"`
}

type projectType struct {
	node       DependencyNode
	visibility string
}

// DependenciesFromResult adapts the batch call inspection result without a
// second source walk.
func DependenciesFromResult(input *calls.Result, request DependencyRequest) (DependencyResult, error) {
	if input == nil {
		return DependencyResult{}, fmt.Errorf("dependency graph input is nil")
	}
	return Dependencies(snapshotFromResult(input), request), nil
}

// Dependencies projects a resolved call graph into procedure, module, and
// project-type relationships. It consumes the same Snapshot produced by both
// batch inspection and the incremental LSP workspace index.
func Dependencies(input Snapshot, request DependencyRequest) DependencyResult {
	nodes := map[string]DependencyNode{}
	procedureByKey := map[string]string{}
	procedureByCandidate := map[string]string{}
	typesByName := map[string][]projectType{}

	for _, sym := range input.Symbols {
		module := moduleNodeFromSymbol(sym)
		if _, exists := nodes[module.ID]; module.ID != "" && (!exists || strings.EqualFold(sym.Kind, "module")) {
			nodes[module.ID] = module
		}
		if procedureKind(sym.Kind) {
			node := procedureNodeFromSymbol(sym)
			nodes[node.ID] = node
			procedureByKey[procedureSymbolKey(sym)] = node.ID
			procedureByCandidate[candidateKey(sym.Module+"."+sym.Name, sym.Kind, sym.File, sym.Line)] = node.ID
		}
		if typ, ok := typeNodeFromSymbol(sym); ok {
			nodes[typ.ID] = typ
			typesByName[strings.ToLower(typ.Name)] = append(typesByName[strings.ToLower(typ.Name)], projectType{node: typ, visibility: sym.Visibility})
		}
	}
	for name := range typesByName {
		sort.Slice(typesByName[name], func(i, j int) bool { return typesByName[name][i].node.ID < typesByName[name][j].node.ID })
	}

	edges := map[string]DependencyEdge{}
	uncertain := make([]UncertainEdge, 0)
	for _, call := range input.Calls {
		if isNewCall(call) {
			// inspect calls retains New expressions for backwards compatibility;
			// their typed dependency is represented by TypeReferences below.
			continue
		}
		callerID, callerOK := procedureByKey[callSiteProcedureKey(call)]
		if call.Resolution.Status == "matched" && len(call.Resolution.Candidates) == 1 && callerOK {
			candidate := call.Resolution.Candidates[0]
			calleeID, calleeOK := procedureByCandidate[candidateKey(candidate.QualifiedName, candidate.Kind, candidate.File, candidate.Line)]
			if calleeOK {
				evidence := evidenceForCall("call", call, callerID, calleeID)
				addEdge(edges, "calls", callerID, calleeID, evidence)
				addEdge(edges, "calls", moduleID(call.Module, call.File), moduleID(nodes[calleeID].Module, nodes[calleeID].File), evidence)
				continue
			}
		}
		from := moduleID(call.Module, call.File)
		if callerOK {
			from = callerID
		}
		uncertain = append(uncertain, UncertainEdge{
			Status: call.Resolution.Status, From: from, Callee: call.Callee.Text,
			Candidates: append([]calls.Candidate(nil), call.Resolution.Candidates...),
			Evidence:   evidenceForCall("call", call, "", ""),
		})
	}

	for _, ref := range input.TypeReferences {
		from := moduleID(ref.Module, ref.File)
		if target, status, ok := resolveType(typesByName, ref.Target, ref.Module); ok {
			addEdge(edges, ref.Kind, from, target.ID, evidenceForTypeReference(ref))
		} else if status != "" {
			uncertain = append(uncertain, UncertainEdge{Status: status, From: from, Callee: ref.Target, Evidence: evidenceForTypeReference(ref)})
		}
	}

	result := DependencyResult{Target: "dependencies", Nodes: sortedDependencyNodes(nodes), Edges: sortedDependencyEdges(edges), UncertainEdges: uncertain}
	sort.Slice(result.UncertainEdges, func(i, j int) bool {
		return uncertainKey(result.UncertainEdges[i]) < uncertainKey(result.UncertainEdges[j])
	})
	return filterDependencies(result, request.Module)
}

func moduleNodeFromSymbol(sym Symbol) DependencyNode {
	if sym.Module == "" {
		return DependencyNode{}
	}
	return DependencyNode{ID: moduleID(sym.Module, sym.File), Kind: "module", Name: sym.Module, Module: sym.Module, ModuleKind: sym.ModuleKind, File: sym.File, Location: symbolLocation(sym)}
}

func procedureNodeFromSymbol(sym Symbol) DependencyNode {
	id := ID{Module: sym.Module, QualifiedName: sym.Module + "." + sym.Name, Kind: sym.Kind, File: sym.File, Line: sym.Line, Column: sym.Column}
	return DependencyNode{ID: "procedure|" + id.String(), Kind: "procedure", Name: sym.Name, Module: sym.Module, ModuleKind: sym.ModuleKind, File: sym.File, Location: symbolLocation(sym)}
}

func typeNodeFromSymbol(sym Symbol) (DependencyNode, bool) {
	if strings.EqualFold(sym.Kind, "module") && strings.EqualFold(sym.ModuleKind, "class") {
		return DependencyNode{ID: fmt.Sprintf("type|class|%s|%s", strings.ToLower(sym.Module), sym.File), Kind: "type", Name: sym.Module, Module: sym.Module, ModuleKind: sym.ModuleKind, File: sym.File, Location: symbolLocation(sym)}, true
	}
	if !strings.EqualFold(sym.Kind, "type") && !strings.EqualFold(sym.Kind, "enum") {
		return DependencyNode{}, false
	}
	return DependencyNode{ID: fmt.Sprintf("type|%s|%s|%s|%s|%d", strings.ToLower(sym.Kind), strings.ToLower(sym.Module), strings.ToLower(sym.Name), sym.File, sym.Line), Kind: "type", Name: sym.Name, Module: sym.Module, ModuleKind: sym.ModuleKind, File: sym.File, Location: symbolLocation(sym)}, true
}

func moduleID(module, file string) string {
	return fmt.Sprintf("module|%s|%s", strings.ToLower(strings.TrimSpace(module)), file)
}

func procedureSymbolKey(sym Symbol) string {
	return callerKey(sym.Module+"."+sym.Name, sym.Kind, sym.File)
}
func callSiteProcedureKey(call calls.Call) string {
	if call.Caller == nil {
		return ""
	}
	return callerKey(call.Caller.QualifiedName, call.Caller.Kind, call.File)
}
func callerKey(qualifiedName, kind, file string) string {
	return strings.ToLower(qualifiedName) + "|" + strings.ToLower(kind) + "|" + file
}
func candidateKey(qualifiedName, kind, file string, line int) string {
	return strings.ToLower(qualifiedName) + "|" + strings.ToLower(kind) + "|" + file + fmt.Sprintf("|%d", line)
}

func resolveType(byName map[string][]projectType, raw, callerModule string) (DependencyNode, string, bool) {
	qualifier, name := typeReferenceName(raw)
	name = strings.ToLower(name)
	candidates := byName[name]
	if len(candidates) == 0 {
		return DependencyNode{}, "", false
	}
	visible := make([]projectType, 0, len(candidates))
	for _, candidate := range candidates {
		if qualifier != "" && !strings.EqualFold(candidate.node.Module, qualifier) {
			continue
		}
		if strings.EqualFold(candidate.visibility, "private") && !strings.EqualFold(candidate.node.Module, callerModule) {
			continue
		}
		visible = append(visible, candidate)
	}
	if len(visible) != 1 {
		if len(visible) == 0 {
			if qualifier != "" {
				return DependencyNode{}, "external", false
			}
			return DependencyNode{}, "unresolved", false
		}
		return DependencyNode{}, "ambiguous", false
	}
	return visible[0].node, "", true
}

func typeReferenceName(raw string) (qualifier, name string) {
	raw = strings.TrimSpace(raw)
	raw = strings.TrimPrefix(strings.TrimSpace(raw), "New ")
	raw = strings.Trim(raw, "[]() ")
	parts := strings.Split(raw, ".")
	if len(parts) == 0 {
		return "", ""
	}
	name = strings.Trim(strings.TrimSpace(parts[len(parts)-1]), "[]() ")
	if len(parts) == 2 {
		qualifier = strings.Trim(strings.TrimSpace(parts[0]), "[]() ")
	}
	return qualifier, name
}

func isNewCall(call calls.Call) bool {
	return strings.HasPrefix(strings.ToLower(strings.TrimSpace(call.Callee.Text)), "new ")
}

func symbolLocation(sym Symbol) Location {
	return Location{File: sym.File, StartLine: sym.Line, StartColumn: sym.Column, EndLine: sym.EndLine, EndColumn: sym.EndColumn}
}
func evidenceForTypeReference(ref calls.TypeReference) DependencyEvidence {
	return DependencyEvidence{Kind: ref.Kind, File: ref.File, Location: Location{File: ref.File, StartLine: ref.Range.StartLine, StartColumn: ref.Range.StartColumn, EndLine: ref.Range.EndLine, EndColumn: ref.Range.EndColumn}}
}
func evidenceForCall(kind string, call calls.Call, source, target string) DependencyEvidence {
	return DependencyEvidence{Kind: kind, File: call.File, Location: Location{File: call.File, StartLine: call.Range.StartLine, StartColumn: call.Range.StartColumn, EndLine: call.Range.EndLine, EndColumn: call.Range.EndColumn}, Source: source, Target: target}
}

func addEdge(edges map[string]DependencyEdge, kind, from, to string, evidence DependencyEvidence) {
	if from == "" || to == "" {
		return
	}
	key := kind + "|" + from + "|" + to
	edge := edges[key]
	if edge.Kind == "" {
		edge = DependencyEdge{Kind: kind, From: from, To: to, Evidence: []DependencyEvidence{}}
	}
	for _, existing := range edge.Evidence {
		if evidenceKey(existing) == evidenceKey(evidence) {
			return
		}
	}
	edge.Evidence = append(edge.Evidence, evidence)
	sort.Slice(edge.Evidence, func(i, j int) bool { return evidenceKey(edge.Evidence[i]) < evidenceKey(edge.Evidence[j]) })
	edges[key] = edge
}

func evidenceKey(e DependencyEvidence) string {
	return fmt.Sprintf("%s|%s|%d|%d|%s|%s", e.Kind, e.File, e.Location.StartLine, e.Location.StartColumn, e.Source, e.Target)
}
func uncertainKey(edge UncertainEdge) string {
	return edge.Status + "|" + edge.From + "|" + edge.Callee + "|" + evidenceKey(edge.Evidence)
}

func sortedDependencyNodes(nodes map[string]DependencyNode) []DependencyNode {
	out := make([]DependencyNode, 0, len(nodes))
	for _, node := range nodes {
		out = append(out, node)
	}
	sort.Slice(out, func(i, j int) bool { return out[i].ID < out[j].ID })
	return out
}
func sortedDependencyEdges(edges map[string]DependencyEdge) []DependencyEdge {
	out := make([]DependencyEdge, 0, len(edges))
	for _, edge := range edges {
		out = append(out, edge)
	}
	sort.Slice(out, func(i, j int) bool { return edgeKeyForDependency(out[i]) < edgeKeyForDependency(out[j]) })
	return out
}
func edgeKeyForDependency(edge DependencyEdge) string {
	return edge.Kind + "|" + edge.From + "|" + edge.To
}

func filterDependencies(input DependencyResult, module string) DependencyResult {
	module = strings.TrimSpace(module)
	if module == "" {
		return input
	}
	selected := map[string]bool{}
	for _, node := range input.Nodes {
		if strings.EqualFold(node.Module, module) {
			selected[node.ID] = true
		}
	}
	if len(selected) == 0 {
		return DependencyResult{Target: "dependencies", Nodes: []DependencyNode{}, Edges: []DependencyEdge{}, UncertainEdges: []UncertainEdge{}}
	}
	edges := make([]DependencyEdge, 0)
	sources := make(map[string]bool, len(selected))
	for id := range selected {
		sources[id] = true
	}
	for _, edge := range input.Edges {
		if sources[edge.From] {
			edges = append(edges, edge)
			selected[edge.To] = true
		}
	}
	uncertain := make([]UncertainEdge, 0)
	for _, edge := range input.UncertainEdges {
		if sources[edge.From] {
			uncertain = append(uncertain, edge)
		}
	}
	nodes := make([]DependencyNode, 0, len(selected))
	for _, node := range input.Nodes {
		if selected[node.ID] {
			nodes = append(nodes, node)
		}
	}
	return DependencyResult{Target: "dependencies", Nodes: nodes, Edges: edges, UncertainEdges: uncertain}
}
