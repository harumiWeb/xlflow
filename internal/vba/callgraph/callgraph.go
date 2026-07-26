// Package callgraph turns conservatively resolved VBA call sites into a
// deterministic, project-local procedure graph.
package callgraph

import (
	"errors"
	"fmt"
	"path/filepath"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/calls"
)

const (
	DirectionCallers = "callers"
	DirectionCallees = "callees"
	DirectionBoth    = "both"
)

var ErrTargetNotFound = errors.New("impact target not found")

// ID is a stable identity for one declared project-local procedure.
type ID struct {
	Module        string `json:"module"`
	QualifiedName string `json:"qualified_name"`
	Kind          string `json:"kind"`
	File          string `json:"file"`
	Line          int    `json:"line"`
	Column        int    `json:"column"`
}

func (id ID) String() string {
	return fmt.Sprintf("%s|%s|%s|%s|%d|%d", strings.ToLower(id.Module), strings.ToLower(id.QualifiedName), id.Kind, id.File, id.Line, id.Column)
}

// Node exposes the symbol and declaration location used by impact consumers.
type Node struct {
	ID         ID     `json:"id"`
	Name       string `json:"name"`
	ModuleKind string `json:"module_kind"`
}

// Location identifies a source call site. Lines and columns follow the
// existing tree-sitter source ranges.
type Location struct {
	File        string `json:"file"`
	StartLine   int    `json:"start_line"`
	StartColumn int    `json:"start_column"`
	EndLine     int    `json:"end_line"`
	EndColumn   int    `json:"end_column"`
}

type Edge struct {
	Caller   ID       `json:"caller"`
	Callee   ID       `json:"callee"`
	Location Location `json:"location"`
}

type Cycle struct {
	Nodes []ID `json:"nodes"`
}

type Traversal struct {
	Direction string `json:"direction"`
	Depth     int    `json:"depth"`
}

type Uncertainty struct {
	Ambiguous   int `json:"ambiguous"`
	Unresolved  int `json:"unresolved"`
	External    int `json:"external"`
	BuiltinLike int `json:"builtin_like"`
	MemberCall  int `json:"member_call"`
}

type Result struct {
	Target            Node        `json:"target"`
	Traversal         Traversal   `json:"traversal"`
	Nodes             []Node      `json:"nodes"`
	Edges             []Edge      `json:"edges"`
	DirectCallers     []Node      `json:"direct_callers"`
	DirectCallees     []Node      `json:"direct_callees"`
	UpstreamCallers   []Node      `json:"upstream_callers"`
	DownstreamCallees []Node      `json:"downstream_callees"`
	AffectedModules   []Module    `json:"affected_modules"`
	Cycles            []Cycle     `json:"cycles"`
	Uncertainty       Uncertainty `json:"uncertainty"`
}

type Module struct {
	Name string `json:"name"`
	Kind string `json:"kind"`
	File string `json:"file"`
}

type Request struct {
	Target    string
	Direction string
	Depth     int
}

// Symbol is the protocol-neutral declaration data needed to create a stable
// graph node. It deliberately avoids LSP and tree-sitter-specific types.
type Symbol struct {
	Name       string
	Kind       string
	Module     string
	ModuleKind string
	File       string
	Line       int
	Column     int
}

// Snapshot is one coherent set of declarations and resolved call sites. Both
// CLI batch inspection and the LSP's incremental workspace index can produce
// this value without exposing their storage implementation to graph traversal.
type Snapshot struct {
	Symbols []Symbol
	Calls   []calls.Call
}

// AmbiguousTargetError gives callers the exact stable candidates rather than
// silently combining unrelated same-named procedures.
type AmbiguousTargetError struct{ Candidates []Node }

func (e *AmbiguousTargetError) Error() string { return "impact target is ambiguous" }

type graph struct {
	nodes       map[string]Node
	byQualified map[string][]string
	out         map[string][]Edge
	in          map[string][]Edge
	uncertain   map[string][]calls.Call
}

func Analyze(input *calls.Result, request Request) (Result, error) {
	if input == nil {
		return Result{}, errors.New("call graph input is nil")
	}
	snapshot := Snapshot{Calls: input.Calls, Symbols: make([]Symbol, 0, len(input.Symbols))}
	for _, sym := range input.Symbols {
		kind := input.ModuleKinds[sym.File]
		if kind == "" {
			kind = moduleKindForFile(sym.File)
		}
		snapshot.Symbols = append(snapshot.Symbols, Symbol{Name: sym.Name, Kind: sym.Kind, Module: sym.Module, File: sym.File, Line: sym.StartLine, Column: sym.StartColumn, ModuleKind: kind})
	}
	return AnalyzeSnapshot(snapshot, request)
}

func AnalyzeSnapshot(input Snapshot, request Request) (Result, error) {
	direction := strings.ToLower(strings.TrimSpace(request.Direction))
	if direction == "" {
		direction = DirectionBoth
	}
	if direction != DirectionCallers && direction != DirectionCallees && direction != DirectionBoth {
		return Result{}, fmt.Errorf("unsupported direction %q; expected callers, callees, or both", request.Direction)
	}
	if request.Depth < 0 {
		return Result{}, errors.New("depth must be greater than or equal to 0")
	}
	g := build(input)
	targets := g.byQualified[strings.ToLower(strings.TrimSpace(request.Target))]
	if len(targets) == 0 {
		return Result{}, ErrTargetNotFound
	}
	if len(targets) > 1 {
		candidates := nodesForKeys(g, targets)
		return Result{}, &AmbiguousTargetError{Candidates: candidates}
	}
	targetKey := targets[0]
	result := Result{Target: g.nodes[targetKey], Traversal: Traversal{Direction: direction, Depth: request.Depth}, Nodes: []Node{}, Edges: []Edge{}, DirectCallers: []Node{}, DirectCallees: []Node{}, UpstreamCallers: []Node{}, DownstreamCallees: []Node{}, AffectedModules: []Module{}, Cycles: []Cycle{}}
	visited := map[string]bool{targetKey: true}
	edges := map[string]Edge{}
	callerDistances, callerEdges := map[string]int{}, map[string]Edge{}
	calleeDistances, calleeEdges := map[string]int{}, map[string]Edge{}
	if direction == DirectionCallers || direction == DirectionBoth {
		callerDistances, callerEdges = g.walk(targetKey, false, request.Depth)
	}
	if direction == DirectionCallees || direction == DirectionBoth {
		calleeDistances, calleeEdges = g.walk(targetKey, true, request.Depth)
	}
	for key := range callerDistances {
		visited[key] = true
	}
	for key := range calleeDistances {
		visited[key] = true
	}
	for key, edge := range callerEdges {
		edges[key] = edge
	}
	for key, edge := range calleeEdges {
		edges[key] = edge
	}
	for key := range visited {
		result.Nodes = append(result.Nodes, g.nodes[key])
	}
	for _, edge := range edges {
		result.Edges = append(result.Edges, edge)
	}
	sort.Slice(result.Edges, func(i, j int) bool { return edgeKey(result.Edges[i]) < edgeKey(result.Edges[j]) })
	if request.Depth > 0 {
		if direction == DirectionCallers || direction == DirectionBoth {
			for _, edge := range g.in[targetKey] {
				result.DirectCallers = appendUniqueNode(result.DirectCallers, g.nodes[edge.Caller.String()])
			}
		}
		if direction == DirectionCallees || direction == DirectionBoth {
			for _, edge := range g.out[targetKey] {
				result.DirectCallees = appendUniqueNode(result.DirectCallees, g.nodes[edge.Callee.String()])
			}
		}
	}
	for key, distance := range callerDistances {
		if distance > 1 {
			result.UpstreamCallers = append(result.UpstreamCallers, g.nodes[key])
		}
	}
	for key, distance := range calleeDistances {
		if distance > 1 {
			result.DownstreamCallees = append(result.DownstreamCallees, g.nodes[key])
		}
	}
	result.Uncertainty = g.uncertaintyFor(visited)
	result.AffectedModules = modulesFor(result.Nodes)
	result.Cycles = cyclesFor(visited, result.Edges)
	sortNodes(result.Nodes)
	sortNodes(result.DirectCallers)
	sortNodes(result.DirectCallees)
	sortNodes(result.UpstreamCallers)
	sortNodes(result.DownstreamCallees)
	return result, nil
}

func build(input Snapshot) graph {
	g := graph{nodes: map[string]Node{}, byQualified: map[string][]string{}, out: map[string][]Edge{}, in: map[string][]Edge{}, uncertain: map[string][]calls.Call{}}
	moduleKinds := map[string]string{}
	for _, sym := range input.Symbols {
		if sym.Kind == "module" {
			moduleKinds[sym.Module] = sym.ModuleKind
		}
	}
	for _, sym := range input.Symbols {
		if !procedureKind(sym.Kind) {
			continue
		}
		kind := moduleKinds[sym.Module]
		if kind == "" {
			kind = sym.ModuleKind
		}
		if kind == "" {
			kind = moduleKindForFile(sym.File)
		}
		node := nodeFromSymbol(sym, kind)
		key := node.ID.String()
		g.nodes[key] = node
		qualified := strings.ToLower(node.ID.QualifiedName)
		g.byQualified[qualified] = append(g.byQualified[qualified], key)
	}
	for qualified := range g.byQualified {
		sort.Strings(g.byQualified[qualified])
	}
	for _, call := range input.Calls {
		if call.Caller == nil {
			continue
		}
		callerKey, ok := g.callerKey(call)
		if !ok {
			continue
		}
		if call.Resolution.Status != "matched" || len(call.Resolution.Candidates) != 1 {
			g.uncertain[callerKey] = append(g.uncertain[callerKey], call)
			continue
		}
		candidate := call.Resolution.Candidates[0]
		calleeKey := ""
		for key, node := range g.nodes {
			if strings.EqualFold(node.ID.QualifiedName, candidate.QualifiedName) && node.ID.Kind == candidate.Kind && node.ID.File == candidate.File && node.ID.Line == candidate.Line {
				calleeKey = key
				break
			}
		}
		if calleeKey == "" {
			continue
		}
		edge := Edge{Caller: g.nodes[callerKey].ID, Callee: g.nodes[calleeKey].ID, Location: Location{File: call.File, StartLine: call.Range.StartLine, StartColumn: call.Range.StartColumn, EndLine: call.Range.EndLine, EndColumn: call.Range.EndColumn}}
		key := edgeKey(edge)
		if hasEdge(g.out[callerKey], key) {
			continue
		}
		g.out[callerKey] = append(g.out[callerKey], edge)
		g.in[calleeKey] = append(g.in[calleeKey], edge)
	}
	for key := range g.out {
		sort.Slice(g.out[key], func(i, j int) bool { return edgeKey(g.out[key][i]) < edgeKey(g.out[key][j]) })
	}
	for key := range g.in {
		sort.Slice(g.in[key], func(i, j int) bool { return edgeKey(g.in[key][i]) < edgeKey(g.in[key][j]) })
	}
	return g
}

func (g graph) callerKey(call calls.Call) (string, bool) {
	if call.Caller == nil {
		return "", false
	}
	for _, key := range g.byQualified[strings.ToLower(call.Caller.QualifiedName)] {
		node := g.nodes[key]
		if node.ID.Kind == call.Caller.Kind && node.ID.File == call.File {
			return key, true
		}
	}
	return "", false
}

func (g graph) walk(root string, downstream bool, maxDepth int) (map[string]int, map[string]Edge) {
	distances := map[string]int{}
	selected := map[string]Edge{}
	queue := []string{root}
	distances[root] = 0
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		distance := distances[current]
		if distance >= maxDepth {
			continue
		}
		edges := g.out[current]
		if !downstream {
			edges = g.in[current]
		}
		for _, edge := range edges {
			next := edge.Callee.String()
			if !downstream {
				next = edge.Caller.String()
			}
			selected[edgeKey(edge)] = edge
			if _, seen := distances[next]; !seen {
				distances[next] = distance + 1
				queue = append(queue, next)
			}
		}
	}
	delete(distances, root)
	return distances, selected
}

func (g graph) uncertaintyFor(nodes map[string]bool) Uncertainty {
	var out Uncertainty
	for key := range nodes {
		for _, call := range g.uncertain[key] {
			switch call.Resolution.Status {
			case "ambiguous":
				out.Ambiguous++
			case "external":
				out.External++
			case "builtin_like":
				out.BuiltinLike++
			case "member_call":
				out.MemberCall++
			default:
				out.Unresolved++
			}
		}
	}
	return out
}
func procedureKind(kind string) bool {
	switch kind {
	case "sub", "function", "property", "property_get", "property_let", "property_set", "declare", "declare_sub", "declare_function":
		return true
	}
	return false
}
func nodeFromSymbol(sym Symbol, moduleKind string) Node {
	return Node{ID: ID{Module: sym.Module, QualifiedName: sym.Module + "." + sym.Name, Kind: sym.Kind, File: sym.File, Line: sym.Line, Column: sym.Column}, Name: sym.Name, ModuleKind: moduleKind}
}
func moduleKindForFile(file string) string {
	if strings.EqualFold(filepath.Ext(file), ".cls") {
		return "class"
	}
	if strings.EqualFold(filepath.Ext(file), ".frm") {
		return "form"
	}
	return "standard"
}
func edgeKey(edge Edge) string {
	return edge.Caller.String() + "->" + edge.Callee.String() + fmt.Sprintf("@%s:%d:%d", edge.Location.File, edge.Location.StartLine, edge.Location.StartColumn)
}
func hasEdge(edges []Edge, key string) bool {
	for _, edge := range edges {
		if edgeKey(edge) == key {
			return true
		}
	}
	return false
}
func nodesForKeys(g graph, keys []string) []Node {
	result := make([]Node, 0, len(keys))
	for _, key := range keys {
		result = append(result, g.nodes[key])
	}
	sortNodes(result)
	return result
}
func appendUniqueNode(nodes []Node, node Node) []Node {
	for _, existing := range nodes {
		if existing.ID == node.ID {
			return nodes
		}
	}
	return append(nodes, node)
}
func sortNodes(nodes []Node) {
	sort.Slice(nodes, func(i, j int) bool { return nodes[i].ID.String() < nodes[j].ID.String() })
}
func modulesFor(nodes []Node) []Module {
	seen := map[string]Module{}
	for _, node := range nodes {
		key := strings.ToLower(node.ID.Module) + "|" + node.ID.File
		seen[key] = Module{Name: node.ID.Module, Kind: node.ModuleKind, File: node.ID.File}
	}
	out := make([]Module, 0, len(seen))
	for _, module := range seen {
		out = append(out, module)
	}
	sort.Slice(out, func(i, j int) bool { return out[i].Name+"|"+out[i].File < out[j].Name+"|"+out[j].File })
	return out
}
func cyclesFor(visited map[string]bool, edges []Edge) []Cycle {
	adjacency := map[string][]string{}
	sortedEdges := append([]Edge(nil), edges...)
	sort.Slice(sortedEdges, func(i, j int) bool { return edgeKey(sortedEdges[i]) < edgeKey(sortedEdges[j]) })
	for _, edge := range sortedEdges {
		a, b := edge.Caller.String(), edge.Callee.String()
		if visited[a] && visited[b] {
			adjacency[a] = append(adjacency[a], b)
		}
	}
	for node := range adjacency {
		sort.Strings(adjacency[node])
	}
	seen, stack, onStack := map[string]bool{}, []string{}, map[string]int{}
	unique := map[string]Cycle{}
	var visit func(string)
	visit = func(node string) {
		seen[node] = true
		onStack[node] = len(stack)
		stack = append(stack, node)
		for _, next := range adjacency[node] {
			if index, ok := onStack[next]; ok {
				ids := make([]ID, 0, len(stack)-index)
				for _, key := range stack[index:] {
					parts := strings.Split(key, "|")
					_ = parts
					ids = append(ids, idForKey(key, visited, edges))
				}
				key := cycleKey(ids)
				unique[key] = Cycle{Nodes: ids}
				continue
			}
			if !seen[next] {
				visit(next)
			}
		}
		delete(onStack, node)
		stack = stack[:len(stack)-1]
	}
	keys := make([]string, 0, len(visited))
	for key := range visited {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	for _, key := range keys {
		if !seen[key] {
			visit(key)
		}
	}
	out := make([]Cycle, 0, len(unique))
	for _, cycle := range unique {
		out = append(out, cycle)
	}
	sort.Slice(out, func(i, j int) bool { return cycleKey(out[i].Nodes) < cycleKey(out[j].Nodes) })
	return out
}
func idForKey(key string, _ map[string]bool, edges []Edge) ID {
	for _, edge := range edges {
		if edge.Caller.String() == key {
			return edge.Caller
		}
		if edge.Callee.String() == key {
			return edge.Callee
		}
	}
	parts := strings.Split(key, "|")
	return ID{Module: parts[0], QualifiedName: parts[1], Kind: parts[2], File: parts[3]}
}
func cycleKey(ids []ID) string {
	values := make([]string, len(ids))
	for i, id := range ids {
		values[i] = id.String()
	}
	sort.Strings(values)
	return strings.Join(values, ",")
}
