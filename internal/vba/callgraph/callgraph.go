// Package callgraph turns conservatively resolved VBA call sites into a
// deterministic, project-local procedure graph.
package callgraph

import (
	"context"
	"errors"
	"fmt"
	"path/filepath"
	"sort"
	"strings"
	"unicode"
	"unicode/utf8"

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
	Visibility string `json:"-"`
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
	// Edges follows Nodes in traversal order. The i-th edge leaves Nodes[i]
	// and enters Nodes[(i+1)%len(Nodes)]. Parallel call sites between the same
	// endpoints are represented by their earliest source edge.
	Edges []Edge `json:"edges"`
}

// CyclicComponent is one strongly connected component that contains a
// directed cycle. Nodes contains every procedure in the component, while
// Witness contains one deterministic representative cycle. The witness is
// intentionally bounded to one cycle per component; callers that need every
// elementary cycle should continue to use FindCyclesContext.
type CyclicComponent struct {
	Nodes   []ID  `json:"nodes"`
	Witness Cycle `json:"witness"`
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
	EndLine    int
	EndColumn  int
	Parent     string
	Visibility string
	ReturnType string
	Signature  string
}

// Snapshot is one coherent set of declarations and resolved call sites. Both
// CLI batch inspection and the LSP's incremental workspace index can produce
// this value without exposing their storage implementation to graph traversal.
type Snapshot struct {
	Symbols           []Symbol
	Calls             []calls.Call
	TypeReferences    []calls.TypeReference
	DynamicReferences []calls.DynamicReference
}

type RootConfidence string

const (
	RootConfirmed RootConfidence = "confirmed"
	RootPossible  RootConfidence = "possible"
)

type Root struct {
	Target     string
	Confidence RootConfidence
	Reason     string
}

type ReachabilityRequest struct {
	Roots []Root
}

type ReachabilityResult struct {
	Confirmed   []Node
	Possible    []Node
	Unreachable []Node
	Clusters    [][]Node
}

// AmbiguousTargetError gives callers the exact stable candidates rather than
// silently combining unrelated same-named procedures.
type AmbiguousTargetError struct{ Candidates []Node }

func (e *AmbiguousTargetError) Error() string { return "impact target is ambiguous" }

type graph struct {
	nodes       map[string]Node
	byQualified map[string][]string
	byName      map[string][]string
	byCandidate map[procedureIdentity]candidateIndexEntry
	out         map[string][]Edge
	in          map[string][]Edge
	uncertain   map[string][]calls.Call
}

// procedureIdentity contains the fields that a uniquely resolved call
// candidate must match before it can become a confirmed graph edge. The file
// component is normalized for indexing, while build still rechecks the
// original file string below to preserve the historical exact-file match.
type procedureIdentity struct {
	qualifiedName string
	kind          string
	file          string
	line          int
}

// candidateIndexEntry retains every graph node in a normalized identity
// bucket. Most buckets contain one node and therefore avoid a slice
// allocation; alternates are retained only for normalized collisions and are
// checked against the original identity fields during lookup.
type candidateIndexEntry struct {
	first      string
	alternates []string
}

func Analyze(input *calls.Result, request Request) (Result, error) {
	if input == nil {
		return Result{}, errors.New("call graph input is nil")
	}
	return AnalyzeSnapshot(snapshotFromResult(input), request)
}

func snapshotFromResult(input *calls.Result) Snapshot {
	snapshot := Snapshot{Calls: input.Calls, TypeReferences: input.TypeReferences, DynamicReferences: input.DynamicReferences, Symbols: make([]Symbol, 0, len(input.Symbols))}
	for _, sym := range input.Symbols {
		kind := input.ModuleKinds[sym.File]
		if kind == "" {
			kind = moduleKindForFile(sym.File)
		}
		snapshot.Symbols = append(snapshot.Symbols, Symbol{
			Name: sym.Name, Kind: sym.Kind, Module: sym.Module, File: sym.File,
			Line: sym.StartLine, Column: sym.StartColumn, EndLine: sym.EndLine, EndColumn: sym.EndColumn,
			ModuleKind: kind, Parent: sym.Parent, Visibility: sym.Visibility, ReturnType: sym.ReturnType, Signature: sym.Signature,
		})
	}
	return snapshot
}

// SnapshotFromResult adapts the shared calls result for graph consumers.
func SnapshotFromResult(input *calls.Result) Snapshot {
	if input == nil {
		return Snapshot{}
	}
	return snapshotFromResult(input)
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
	g := graph{nodes: map[string]Node{}, byQualified: map[string][]string{}, byName: map[string][]string{}, byCandidate: map[procedureIdentity]candidateIndexEntry{}, out: map[string][]Edge{}, in: map[string][]Edge{}, uncertain: map[string][]calls.Call{}}
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
		g.byName[strings.ToLower(node.Name)] = append(g.byName[strings.ToLower(node.Name)], key)
	}
	// The candidate index is populated immediately below and has one entry per
	// procedure in the common case. Reserve that capacity after the node pass so
	// map growth does not become part of the matched-call hot path.
	g.byCandidate = make(map[procedureIdentity]candidateIndexEntry, len(g.nodes))
	for qualified := range g.byQualified {
		sort.Strings(g.byQualified[qualified])
	}
	for name := range g.byName {
		sort.Strings(g.byName[name])
	}
	// Build the candidate identity index once for this graph snapshot. Keep all
	// normalized collisions so an exact raw-file match can still be selected.
	for key, node := range g.nodes {
		identity := procedureIdentityForNode(node)
		entry, exists := g.byCandidate[identity]
		if !exists {
			entry.first = key
		} else {
			entry.alternates = append(entry.alternates, key)
		}
		g.byCandidate[identity] = entry
	}
	seenEdges := make(map[string]struct{}, len(input.Calls))
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
		calleeKey := g.candidateKey(candidate)
		if calleeKey == "" {
			continue
		}
		edge := Edge{Caller: g.nodes[callerKey].ID, Callee: g.nodes[calleeKey].ID, Location: Location{File: call.File, StartLine: call.Range.StartLine, StartColumn: call.Range.StartColumn, EndLine: call.Range.EndLine, EndColumn: call.Range.EndColumn}}
		key := edgeKey(edge)
		if _, exists := seenEdges[key]; exists {
			continue
		}
		seenEdges[key] = struct{}{}
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

func procedureIdentityForNode(node Node) procedureIdentity {
	return procedureIdentity{
		qualifiedName: foldEqualFold(node.ID.QualifiedName),
		kind:          node.ID.Kind,
		file:          normalizeProcedureFile(node.ID.File),
		line:          node.ID.Line,
	}
}

func procedureIdentityForCandidate(candidate calls.Candidate) procedureIdentity {
	return procedureIdentity{
		qualifiedName: foldEqualFold(candidate.QualifiedName),
		kind:          candidate.Kind,
		file:          normalizeProcedureFile(candidate.File),
		line:          candidate.Line,
	}
}

// candidateKey narrows a candidate to its canonical identity bucket, then
// rechecks the original comparison used by the graph builder. Normalized path
// collisions may contain a valid exact match, while multiple exact matches
// remain ambiguous and must not become a confirmed edge.
func (g graph) candidateKey(candidate calls.Candidate) string {
	identity := procedureIdentityForCandidate(candidate)
	entry, ok := g.byCandidate[identity]
	if !ok {
		return ""
	}
	match := ""
	if g.candidateNodeMatches(entry.first, candidate) {
		match = entry.first
	}
	for _, key := range entry.alternates {
		if !g.candidateNodeMatches(key, candidate) {
			continue
		}
		if match != "" {
			return ""
		}
		match = key
	}
	return match
}

func (g graph) candidateNodeMatches(key string, candidate calls.Candidate) bool {
	node, ok := g.nodes[key]
	return ok && strings.EqualFold(node.ID.QualifiedName, candidate.QualifiedName) &&
		node.ID.Kind == candidate.Kind && node.ID.File == candidate.File && node.ID.Line == candidate.Line
}

func normalizeProcedureFile(file string) string {
	if strings.TrimSpace(file) == "" {
		return ""
	}
	return filepath.ToSlash(filepath.Clean(file))
}

// foldEqualFold returns a canonical key for Go's Unicode simple-fold
// equivalence. strings.ToLower is insufficient for a few simple-fold cycles
// (for example, Kelvin sign), so canonicalize each rune through its complete
// unicode.SimpleFold cycle. This keeps indexing lossless while retaining the
// final strings.EqualFold guard in candidateKey.
func foldEqualFold(value string) string {
	ascii := true
	needsUpper := false
	for i := 0; i < len(value); i++ {
		if value[i] >= utf8.RuneSelf {
			ascii = false
			break
		}
		if value[i] >= 'a' && value[i] <= 'z' {
			needsUpper = true
		}
	}
	if ascii {
		if !needsUpper {
			return value
		}
		folded := []byte(value)
		for i, char := range folded {
			if char >= 'a' && char <= 'z' {
				folded[i] -= 'a' - 'A'
			}
		}
		return string(folded)
	}
	var folded strings.Builder
	folded.Grow(len(value))
	for _, r := range value {
		canonical := r
		for next := unicode.SimpleFold(r); next != r; next = unicode.SimpleFold(next) {
			if next < canonical {
				canonical = next
			}
		}
		folded.WriteRune(canonical)
	}
	return folded.String()
}

// AnalyzeReachability walks the confirmed project call graph from explicit
// roots and separately propagates targets discovered through dynamic callback
// references. Dynamic dispatch never becomes a confirmed graph edge.
func AnalyzeReachability(input Snapshot, request ReachabilityRequest) ReachabilityResult {
	g := build(input)
	confirmed := map[string]bool{}
	possible := map[string]bool{}

	for _, root := range request.Roots {
		keys, exact := g.rootKeysWithExact(root.Target)
		if len(keys) == 0 {
			continue
		}
		if root.Confidence == RootPossible || !exact || len(keys) != 1 {
			for _, key := range keys {
				possible[key] = true
			}
			continue
		}
		confirmed[keys[0]] = true
	}

	refsByCaller := map[string][]calls.DynamicReference{}
	for _, ref := range input.DynamicReferences {
		if key := dynamicCallerKey(g, ref); key != "" {
			refsByCaller[key] = append(refsByCaller[key], ref)
		}
	}

	propagateConfirmed(g, confirmed)
	changed := true
	for changed {
		changed = false
		for key := range confirmed {
			if propagateDynamicReferences(g, refsByCaller[key], confirmed, possible) {
				changed = true
			}
		}
		for key := range possible {
			if confirmed[key] {
				delete(possible, key)
				changed = true
				continue
			}
			for _, edge := range g.out[key] {
				next := edge.Callee.String()
				if !confirmed[next] && !possible[next] {
					possible[next] = true
					changed = true
				}
			}
			if propagateDynamicReferences(g, refsByCaller[key], confirmed, possible) {
				changed = true
			}
		}
		propagateConfirmed(g, confirmed)
	}

	result := ReachabilityResult{Confirmed: []Node{}, Possible: []Node{}, Unreachable: []Node{}, Clusters: [][]Node{}}
	for key, node := range g.nodes {
		switch {
		case confirmed[key]:
			result.Confirmed = append(result.Confirmed, node)
		case possible[key]:
			result.Possible = append(result.Possible, node)
		case strings.EqualFold(node.Visibility, "Private"):
			result.Unreachable = append(result.Unreachable, node)
		}
	}
	sortNodes(result.Confirmed)
	sortNodes(result.Possible)
	sortNodes(result.Unreachable)
	result.Clusters = unreachablePrivateClusters(g, result.Unreachable)
	return result
}

func propagateConfirmed(g graph, confirmed map[string]bool) {
	queue := make([]string, 0, len(confirmed))
	for key := range confirmed {
		queue = append(queue, key)
	}
	for len(queue) > 0 {
		key := queue[0]
		queue = queue[1:]
		for _, edge := range g.out[key] {
			next := edge.Callee.String()
			if !confirmed[next] {
				confirmed[next] = true
				queue = append(queue, next)
			}
		}
	}
}

func propagateDynamicReferences(g graph, refs []calls.DynamicReference, confirmed, possible map[string]bool) bool {
	changed := false
	for _, ref := range refs {
		if ref.Target == "" || ref.Kind != "static" {
			for key := range g.nodes {
				if !confirmed[key] && !possible[key] {
					possible[key] = true
					changed = true
				}
			}
			continue
		}
		keys := g.dynamicKeys(ref.Target)
		for _, key := range keys {
			if !confirmed[key] && !possible[key] {
				possible[key] = true
				changed = true
			}
		}
	}
	return changed
}

func dynamicCallerKey(g graph, ref calls.DynamicReference) string {
	if ref.Caller == nil {
		return ""
	}
	key, ok := g.lookupCallerKey(ref.Caller.QualifiedName, ref.Caller.Kind, ref.File)
	if !ok {
		return ""
	}
	return key
}

func (g graph) lookupCallerKey(qualifiedName, kind, file string) (string, bool) {
	for _, key := range g.byQualified[strings.ToLower(qualifiedName)] {
		node := g.nodes[key]
		if node.ID.Kind == kind && node.ID.File == file {
			return key, true
		}
	}
	return "", false
}

func (g graph) rootKeys(target string) []string {
	keys, _ := g.rootKeysWithExact(target)
	return keys
}

func (g graph) rootKeysWithExact(target string) ([]string, bool) {
	target = normalizeDynamicTarget(target)
	if target == "" {
		return nil, false
	}
	if keys := g.byQualified[strings.ToLower(target)]; len(keys) > 0 {
		return append([]string(nil), keys...), true
	}
	if index := strings.LastIndex(target, "."); index >= 0 {
		if keys := g.byName[strings.ToLower(target[index+1:])]; len(keys) > 0 {
			return append([]string(nil), keys...), false
		}
	}
	return append([]string(nil), g.byName[strings.ToLower(target)]...), true
}

func (g graph) dynamicKeys(target string) []string {
	return g.rootKeys(target)
}

func normalizeDynamicTarget(target string) string {
	target = strings.TrimSpace(target)
	if index := strings.LastIndex(target, "!"); index >= 0 {
		target = target[index+1:]
	}
	target = strings.Trim(strings.TrimSpace(target), "'")
	return strings.TrimSpace(target)
}

func unreachablePrivateClusters(g graph, nodes []Node) [][]Node {
	keys := map[string]bool{}
	for _, node := range nodes {
		keys[node.ID.String()] = true
	}
	seen := map[string]bool{}
	clusters := make([][]Node, 0)
	for _, node := range nodes {
		start := node.ID.String()
		if seen[start] {
			continue
		}
		seen[start] = true
		queue := []string{start}
		cluster := []Node{}
		for len(queue) > 0 {
			key := queue[0]
			queue = queue[1:]
			cluster = append(cluster, g.nodes[key])
			for _, edge := range append(append([]Edge{}, g.out[key]...), g.in[key]...) {
				next := edge.Callee.String()
				if edge.Caller.String() != key {
					next = edge.Caller.String()
				}
				if keys[next] && !seen[next] {
					seen[next] = true
					queue = append(queue, next)
				}
			}
		}
		sortNodes(cluster)
		clusters = append(clusters, cluster)
	}
	sort.Slice(clusters, func(i, j int) bool {
		return clusters[i][0].ID.String() < clusters[j][0].ID.String()
	})
	return clusters
}

func (g graph) callerKey(call calls.Call) (string, bool) {
	if call.Caller == nil {
		return "", false
	}
	return g.lookupCallerKey(call.Caller.QualifiedName, call.Caller.Kind, call.File)
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
	case "sub", "function", "property", "property_get", "property_let", "property_set", "declare", "declare_sub", "declare_function", "event":
		return true
	}
	return false
}
func nodeFromSymbol(sym Symbol, moduleKind string) Node {
	return Node{ID: ID{Module: sym.Module, QualifiedName: sym.Module + "." + sym.Name, Kind: sym.Kind, File: sym.File, Line: sym.Line, Column: sym.Column}, Name: sym.Name, ModuleKind: moduleKind, Visibility: sym.Visibility}
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

// FindCycles returns every elementary directed cycle in the confirmed,
// project-local call graph represented by input. Cycles are graph-wide (they
// are not restricted to a traversal target), deterministic, and represented
// in canonical rotation order. Unresolved, ambiguous, external, built-in-like,
// and member calls are intentionally absent because build only adds uniquely
// matched project-local edges.
func FindCycles(input Snapshot) []Cycle {
	cycles, err := FindCyclesContext(context.Background(), input)
	if err != nil {
		// A background context cannot be canceled. Keep the non-context API
		// convenient while retaining cancellation for large workspaces.
		return []Cycle{}
	}
	return cycles
}

// FindCyclesContext is the cancellation-aware form of FindCycles. It returns
// context.Canceled/context.DeadlineExceeded when cancellation is observed and
// never returns a partial cycle set.
func FindCyclesContext(ctx context.Context, input Snapshot) ([]Cycle, error) {
	if ctx == nil {
		ctx = context.Background()
	}
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	g := build(input)
	nodes := make(map[string]ID, len(g.nodes))
	for key, node := range g.nodes {
		nodes[key] = node.ID
	}
	allEdges := make([]Edge, 0)
	for _, edges := range g.out {
		allEdges = append(allEdges, edges...)
	}
	cycles, err := enumerateCycles(ctx, nodes, allEdges, nil)
	if err != nil {
		return nil, err
	}
	return cycles, nil
}

// FindCyclicComponents returns one bounded, deterministic cycle witness for
// every cyclic strongly connected component in input. Unlike FindCycles, it
// does not enumerate elementary cycles and therefore has bounded output per
// component.
func FindCyclicComponents(input Snapshot) []CyclicComponent {
	components, err := FindCyclicComponentsContext(context.Background(), input)
	if err != nil {
		// A background context cannot be canceled. Keep the non-context API
		// convenient while retaining cancellation for large workspaces.
		return []CyclicComponent{}
	}
	return components
}

// FindCyclicComponentsContext is the cancellation-aware bounded cycle
// detector. SCC discovery and representative witness construction are linear
// in the number of graph vertices and unique endpoint edges. It returns
// context.Canceled/context.DeadlineExceeded when cancellation is observed and
// never returns a partial component set.
func FindCyclicComponentsContext(ctx context.Context, input Snapshot) ([]CyclicComponent, error) {
	if ctx == nil {
		ctx = context.Background()
	}
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	return findCyclicComponentsGraphContext(ctx, build(input))
}

// findCyclicComponentsGraphContext runs the bounded detector over an already
// materialized graph. Keeping graph construction outside this helper makes the
// O(V+E) cycle-detection boundary explicit for callers and benchmarks; the
// public Snapshot API still owns normal graph construction.
func findCyclicComponentsGraphContext(ctx context.Context, g graph) ([]CyclicComponent, error) {
	if ctx == nil {
		ctx = context.Background()
	}
	if err := ctx.Err(); err != nil {
		return nil, err
	}

	nodes := make(map[string]ID, len(g.nodes))
	keys := make([]string, 0, len(g.nodes))
	for key, node := range g.nodes {
		nodes[key] = node.ID
		keys = append(keys, key)
	}
	sort.Strings(keys)
	allEdges := make([]Edge, 0)
	for _, edges := range g.out {
		if err := ctx.Err(); err != nil {
			return nil, err
		}
		allEdges = append(allEdges, edges...)
	}
	adjacency, err := buildCycleAdjacency(ctx, nodes, allEdges, nil)
	if err != nil {
		return nil, err
	}

	components, err := stronglyConnectedComponents(ctx, keys, adjacency)
	if err != nil {
		return nil, err
	}
	result := make([]CyclicComponent, 0, len(components))
	for _, component := range components {
		if err := ctx.Err(); err != nil {
			return nil, err
		}
		if !componentContainsCycle(component, adjacency) {
			continue
		}
		witness, err := representativeCycle(ctx, component, adjacency, nodes)
		if err != nil {
			return nil, err
		}
		componentNodes := make([]ID, len(component))
		for index, key := range component {
			componentNodes[index] = nodes[key]
		}
		result = append(result, CyclicComponent{Nodes: componentNodes, Witness: witness})
	}
	return result, nil
}

// representativeCycle chooses the lexicographically first edge leaving the
// component's lexicographically smallest node, then finds a shortest return
// path with deterministic BFS ordering. Since SCCs are disjoint, the total
// work of these searches is O(V+E) after SCC discovery.
func representativeCycle(ctx context.Context, component []string, adjacency map[string][]cycleArc, nodes map[string]ID) (Cycle, error) {
	if len(component) == 0 {
		return Cycle{Nodes: []ID{}, Edges: []Edge{}}, nil
	}
	root := component[0]
	member := make(map[string]bool, len(component))
	for _, key := range component {
		member[key] = true
	}

	var first cycleArc
	found := false
	for _, arc := range adjacency[root] {
		if member[arc.to] {
			first = arc
			found = true
			break
		}
	}
	if !found {
		// This should only be possible for malformed input. SCC classification
		// has already established that a cyclic component has an internal edge.
		return Cycle{Nodes: []ID{}, Edges: []Edge{}}, nil
	}
	if first.to == root {
		return canonicalCycle([]string{root}, []Edge{first.edge}, nodes), nil
	}

	parent := map[string]string{first.to: ""}
	parentEdge := map[string]Edge{}
	queue := []string{first.to}
	foundRoot := false
	for head := 0; head < len(queue) && !foundRoot; head++ {
		if err := ctx.Err(); err != nil {
			return Cycle{}, err
		}
		current := queue[head]
		for _, arc := range adjacency[current] {
			if err := ctx.Err(); err != nil {
				return Cycle{}, err
			}
			if !member[arc.to] {
				continue
			}
			if _, seen := parent[arc.to]; seen {
				continue
			}
			parent[arc.to] = current
			parentEdge[arc.to] = arc.edge
			if arc.to == root {
				foundRoot = true
				break
			}
			queue = append(queue, arc.to)
		}
	}
	if !foundRoot {
		return Cycle{Nodes: []ID{}, Edges: []Edge{}}, nil
	}

	// Walk parent links backwards from root to the first edge's destination.
	// Reversing the collected edges yields the return path destination->root.
	returnEdgesReversed := make([]Edge, 0, len(component))
	for current := root; current != first.to; {
		previous, ok := parent[current]
		if !ok {
			return Cycle{Nodes: []ID{}, Edges: []Edge{}}, nil
		}
		returnEdgesReversed = append(returnEdgesReversed, parentEdge[current])
		current = previous
	}
	returnEdges := make([]Edge, len(returnEdgesReversed))
	for index := range returnEdgesReversed {
		returnEdges[len(returnEdgesReversed)-1-index] = returnEdgesReversed[index]
	}
	cycleKeys := make([]string, 0, len(returnEdges)+1)
	cycleKeys = append(cycleKeys, root, first.to)
	for index := 0; index+1 < len(returnEdges); index++ {
		cycleKeys = append(cycleKeys, returnEdges[index].Callee.String())
	}
	cycleEdges := make([]Edge, 0, len(returnEdges)+1)
	cycleEdges = append(cycleEdges, first.edge)
	cycleEdges = append(cycleEdges, returnEdges...)
	return canonicalCycle(cycleKeys, cycleEdges, nodes), nil
}

func cyclesFor(visited map[string]bool, edges []Edge) []Cycle {
	if len(visited) == 0 || len(edges) == 0 {
		return []Cycle{}
	}
	nodes := make(map[string]ID, len(edges)*2)
	for _, edge := range edges {
		caller, callee := edge.Caller.String(), edge.Callee.String()
		if visited[caller] {
			nodes[caller] = edge.Caller
		}
		if visited[callee] {
			nodes[callee] = edge.Callee
		}
	}
	cycles, err := enumerateCycles(context.Background(), nodes, edges, visited)
	if err != nil {
		return []Cycle{}
	}
	return cycles
}

type cycleArc struct {
	to   string
	edge Edge
}

// buildCycleAdjacency reduces parallel edges to the earliest stable edge for
// each ordered endpoint pair and returns deterministically sorted adjacency.
// The optional allowed set scopes the result for exhaustive inspection while
// preserving the same edge-selection contract used by bounded witnesses.
func buildCycleAdjacency(ctx context.Context, nodes map[string]ID, edges []Edge, allowed map[string]bool) (map[string][]cycleArc, error) {
	if ctx == nil {
		ctx = context.Background()
	}
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	byEndpoint := make(map[string]map[string]Edge)
	for _, edge := range edges {
		if err := ctx.Err(); err != nil {
			return nil, err
		}
		from, to := edge.Caller.String(), edge.Callee.String()
		if allowed != nil && (!allowed[from] || !allowed[to]) {
			continue
		}
		if nodes != nil {
			if _, ok := nodes[from]; !ok {
				continue
			}
			if _, ok := nodes[to]; !ok {
				continue
			}
		}
		if byEndpoint[from] == nil {
			byEndpoint[from] = make(map[string]Edge)
		}
		prior, exists := byEndpoint[from][to]
		if !exists || edgeLess(edge, prior) {
			byEndpoint[from][to] = edge
		}
	}

	fromKeys := make([]string, 0, len(byEndpoint))
	for from := range byEndpoint {
		fromKeys = append(fromKeys, from)
	}
	sort.Strings(fromKeys)
	adjacency := make(map[string][]cycleArc, len(byEndpoint))
	for _, from := range fromKeys {
		if err := ctx.Err(); err != nil {
			return nil, err
		}
		endpointEdges := byEndpoint[from]
		neighbors := make([]string, 0, len(endpointEdges))
		for to := range endpointEdges {
			neighbors = append(neighbors, to)
		}
		sort.Strings(neighbors)
		arcs := make([]cycleArc, 0, len(neighbors))
		for _, to := range neighbors {
			if err := ctx.Err(); err != nil {
				return nil, err
			}
			arcs = append(arcs, cycleArc{to: to, edge: endpointEdges[to]})
		}
		adjacency[from] = arcs
	}
	return adjacency, nil
}

// enumerateCycles uses Johnson's elementary-cycle traversal. Each iteration
// finds the least strongly connected component in the remaining ordered
// subgraph, runs the blocked/unblocked circuit search from that component's
// least node, and then removes that root. This preserves every directed simple
// path while avoiding the target-scoped back-edge omissions of the old DFS.
func enumerateCycles(ctx context.Context, nodes map[string]ID, edges []Edge, allowed map[string]bool) ([]Cycle, error) {
	if ctx == nil {
		ctx = context.Background()
	}
	if err := ctx.Err(); err != nil {
		return nil, err
	}

	adjacency, err := buildCycleAdjacency(ctx, nodes, edges, allowed)
	if err != nil {
		return nil, err
	}

	keysSet := map[string]bool{}
	for key := range nodes {
		if allowed == nil || allowed[key] {
			keysSet[key] = true
		}
	}
	for from, arcs := range adjacency {
		keysSet[from] = true
		for _, arc := range arcs {
			keysSet[arc.to] = true
		}
	}
	keys := make([]string, 0, len(keysSet))
	for key := range keysSet {
		keys = append(keys, key)
	}
	sort.Strings(keys)

	unique := map[string]Cycle{}
	indexByKey := make(map[string]int, len(keys))
	for index, key := range keys {
		indexByKey[key] = index
	}
	for first := 0; first < len(keys); {
		if err := ctx.Err(); err != nil {
			return nil, err
		}
		active := keys[first:]
		components, err := stronglyConnectedComponents(ctx, active, adjacency)
		if err != nil {
			return nil, err
		}
		var component []string
		for _, candidate := range components {
			if !componentContainsCycle(candidate, adjacency) {
				continue
			}
			if component == nil || candidate[0] < component[0] {
				component = candidate
			}
		}
		if component == nil {
			break
		}
		root := component[0]
		componentSet := make(map[string]bool, len(component))
		for _, key := range component {
			componentSet[key] = true
		}
		blocked := make(map[string]bool, len(component))
		blockedBy := make(map[string]map[string]bool, len(component))
		stack := make([]string, 0, len(component))
		stackEdges := make([]Edge, 0, len(component))
		var unblock func(string)
		unblock = func(key string) {
			if !blocked[key] {
				return
			}
			blocked[key] = false
			for predecessor := range blockedBy[key] {
				delete(blockedBy[key], predecessor)
				unblock(predecessor)
			}
		}
		var circuit func(string) (bool, error)
		circuit = func(current string) (bool, error) {
			if err := ctx.Err(); err != nil {
				return false, err
			}
			found := false
			stack = append(stack, current)
			blocked[current] = true
			for _, arc := range adjacency[current] {
				if !componentSet[arc.to] {
					continue
				}
				if arc.to == root {
					nodesInCycle := append([]string(nil), stack...)
					edgesInCycle := append(append([]Edge(nil), stackEdges...), arc.edge)
					cycle := canonicalCycle(nodesInCycle, edgesInCycle, nodes)
					unique[cycleKey(cycle.Nodes)] = cycle
					found = true
					continue
				}
				if blocked[arc.to] {
					continue
				}
				stackEdges = append(stackEdges, arc.edge)
				childFound, err := circuit(arc.to)
				stackEdges = stackEdges[:len(stackEdges)-1]
				if err != nil {
					return false, err
				}
				found = found || childFound
			}
			if found {
				unblock(current)
			} else {
				for _, arc := range adjacency[current] {
					if componentSet[arc.to] && arc.to != root {
						if blockedBy[arc.to] == nil {
							blockedBy[arc.to] = map[string]bool{}
						}
						blockedBy[arc.to][current] = true
					}
				}
			}
			stack = stack[:len(stack)-1]
			return found, nil
		}
		if _, err := circuit(root); err != nil {
			return nil, err
		}
		first = indexByKey[root] + 1
	}

	result := make([]Cycle, 0, len(unique))
	for _, cycle := range unique {
		result = append(result, cycle)
	}
	sort.Slice(result, func(i, j int) bool {
		return cycleKey(result[i].Nodes) < cycleKey(result[j].Nodes)
	})
	return result, nil
}

func stronglyConnectedComponents(ctx context.Context, keys []string, adjacency map[string][]cycleArc) ([][]string, error) {
	allowed := make(map[string]bool, len(keys))
	for _, key := range keys {
		allowed[key] = true
	}
	index := 0
	indices := map[string]int{}
	lowlink := map[string]int{}
	onStack := map[string]bool{}
	stack := []string{}
	components := [][]string{}
	var visit func(string) error
	visit = func(current string) error {
		if err := ctx.Err(); err != nil {
			return err
		}
		indices[current] = index
		lowlink[current] = index
		index++
		stack = append(stack, current)
		onStack[current] = true
		for _, arc := range adjacency[current] {
			if !allowed[arc.to] {
				continue
			}
			if _, seen := indices[arc.to]; !seen {
				if err := visit(arc.to); err != nil {
					return err
				}
				if lowlink[arc.to] < lowlink[current] {
					lowlink[current] = lowlink[arc.to]
				}
			} else if onStack[arc.to] && indices[arc.to] < lowlink[current] {
				lowlink[current] = indices[arc.to]
			}
		}
		if lowlink[current] == indices[current] {
			component := []string{}
			for {
				last := stack[len(stack)-1]
				stack = stack[:len(stack)-1]
				onStack[last] = false
				component = append(component, last)
				if last == current {
					break
				}
			}
			sort.Strings(component)
			components = append(components, component)
		}
		return nil
	}
	for _, key := range keys {
		if _, seen := indices[key]; !seen {
			if err := visit(key); err != nil {
				return nil, err
			}
		}
	}
	sort.Slice(components, func(i, j int) bool { return components[i][0] < components[j][0] })
	return components, nil
}

func componentContainsCycle(component []string, adjacency map[string][]cycleArc) bool {
	if len(component) > 1 {
		return true
	}
	if len(component) == 0 {
		return false
	}
	for _, arc := range adjacency[component[0]] {
		if arc.to == component[0] {
			return true
		}
	}
	return false
}

func canonicalCycle(keys []string, edges []Edge, nodes map[string]ID) Cycle {
	if len(keys) == 0 || len(edges) != len(keys) {
		return Cycle{Nodes: []ID{}, Edges: []Edge{}}
	}
	start := 0
	for index := 1; index < len(keys); index++ {
		if keys[index] < keys[start] {
			start = index
		}
	}
	rotatedKeys := make([]string, len(keys))
	rotatedEdges := make([]Edge, len(edges))
	for index := range keys {
		rotated := (start + index) % len(keys)
		rotatedKeys[index] = keys[rotated]
		rotatedEdges[index] = edges[rotated]
	}
	ids := make([]ID, len(rotatedKeys))
	for index, key := range rotatedKeys {
		if id, ok := nodes[key]; ok {
			ids[index] = id
			continue
		}
		// enumerateCycles always has endpoint IDs available; retain a safe
		// fallback for direct internal callers with an incomplete node map.
		ids[index] = ID{QualifiedName: key}
	}
	return Cycle{Nodes: ids, Edges: rotatedEdges}
}

func edgeLess(a, b Edge) bool {
	la, lb := a.Location, b.Location
	if la.File != lb.File {
		return la.File < lb.File
	}
	if la.StartLine != lb.StartLine {
		return la.StartLine < lb.StartLine
	}
	if la.StartColumn != lb.StartColumn {
		return la.StartColumn < lb.StartColumn
	}
	if la.EndLine != lb.EndLine {
		return la.EndLine < lb.EndLine
	}
	if la.EndColumn != lb.EndColumn {
		return la.EndColumn < lb.EndColumn
	}
	return edgeKey(a) < edgeKey(b)
}

func cycleKey(ids []ID) string {
	values := make([]string, len(ids))
	for i, id := range ids {
		values[i] = id.String()
	}
	return strings.Join(values, ",")
}
