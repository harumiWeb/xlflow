package analyze

import (
	"regexp"
	"strings"

	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// VBA219 deliberately starts with syntax that establishes both a concrete
// acquisition and a procedure-local owner. More dynamic object models belong
// to later resource families.
var (
	resourceWorkbookOpenRe = regexp.MustCompile(`(?i)^\s*set\s+(\[[^\]]+\]|[a-z_][a-z0-9_]*)\s*=\s*(?:application\s*\.\s*)?workbooks\s*\.\s*open\b`)
	resourceFileOpenRe     = regexp.MustCompile(`(?is)^\s*open\b.*\bfor\s+(?:append|binary|input|output|random)\b.*\bas\s*#\s*(\d+|\[[^\]]+\]|[a-z_][a-z0-9_]*)\b`)
	resourceCloseMethodRe  = regexp.MustCompile(`(?i)^\s*(?:call\s+)?(\[[^\]]+\]|[a-z_][a-z0-9_]*)\s*\.\s*close\b`)
	resourceCloseRe        = regexp.MustCompile(`(?is)^\s*close(?:\s+(.*?))?\s*$`)
	resourceHandleRe       = regexp.MustCompile(`(?i)^\s*#\s*(\d+|\[[^\]]+\]|[a-z_][a-z0-9_]*)\s*$`)
	resourceSetAliasRe     = regexp.MustCompile(`(?i)^\s*set\s+(\[[^\]]+\]|[a-z_][a-z0-9_]*)\s*=\s*(\[[^\]]+\]|[a-z_][a-z0-9_]*)\s*$`)
	resourceValueAliasRe   = regexp.MustCompile(`(?i)^\s*(\[[^\]]+\]|[a-z_][a-z0-9_]*)\s*=\s*(\[[^\]]+\]|[a-z_][a-z0-9_]*)\s*$`)
	resourceSetTargetRe    = regexp.MustCompile(`(?i)^\s*set\s+(\[[^\]]+\]|[a-z_][a-z0-9_]*)\s*=`)
	resourceValueTargetRe  = regexp.MustCompile(`(?i)^\s*(\[[^\]]+\]|[a-z_][a-z0-9_]*)\s*=`)
)

type resourceKind string

const (
	resourceWorkbook   resourceKind = "workbook"
	resourceFileHandle resourceKind = "file handle"
)

type resourceAcquisition struct {
	StatementID int
	Kind        resourceKind
	Owner       string
	Line        int
}

type resourceExitWitness struct {
	Kind string
	Line int
}

func (w resourceExitWitness) description() string {
	if w.Line > 0 {
		return w.Kind + " at line " + strconvItoa(w.Line)
	}
	return w.Kind
}

func (a Analyzer) resourceLeakFindings(file parsedFile, proc sourceProcedure) []Finding {
	if proc.Graph == nil {
		return nil
	}
	var findings []Finding
	for _, acquisition := range resourceAcquisitions(proc) {
		witness, leaked := resourceLeakWitness(proc, acquisition)
		if !leaked {
			continue
		}
		resource := "Workbook"
		if acquisition.Kind == resourceFileHandle {
			resource = "VBA file handle"
		}
		findings = append(findings, a.simpleFinding(
			file, proc, acquisition.Line, "VBA219", "warning",
			resource+" acquired here can reach "+witness.Kind+" without Close.",
			"The "+string(acquisition.Kind)+" acquired at line "+strconvItoa(acquisition.Line)+" can leave this procedure through "+witness.description()+" without a matching Close.",
			"Close the "+string(acquisition.Kind)+" in a cleanup path that every exit reaches.",
		))
	}
	return findings
}

func resourceAcquisitions(proc sourceProcedure) []resourceAcquisition {
	acquisitions := make([]resourceAcquisition, 0)
	for statement := range proc.Statements.All() {
		if statement.Recovered {
			continue
		}
		if target, ok := resourceWorkbookOpenTarget(statement.Text); ok && resourceWorkbookOwner(proc, target) {
			acquisitions = append(acquisitions, resourceAcquisition{
				StatementID: statement.ID, Kind: resourceWorkbook, Owner: target, Line: statement.Range.StartLine,
			})
			continue
		}
		if handle, ok := resourceFileOpenHandle(statement.Text); ok && resourceFileHandleOwner(proc, handle) {
			acquisitions = append(acquisitions, resourceAcquisition{
				StatementID: statement.ID, Kind: resourceFileHandle, Owner: handle, Line: statement.Range.StartLine,
			})
		}
	}
	return acquisitions
}

func resourceLeakWitness(proc sourceProcedure, acquisition resourceAcquisition) (resourceExitWitness, bool) {
	graph := proc.Graph
	block, ok := graph.BlockForStatement(acquisition.StatementID)
	if !ok {
		return resourceExitWitness{}, false
	}
	if !graph.IsReachable(block.ID, vbacfg.EdgeFilter{}) {
		return resourceExitWitness{}, false
	}
	initialAliases := map[string]bool{acquisition.Owner: true}
	if acquisition.Kind == resourceFileHandle {
		initialAliases = resourceFileAliasesBefore(proc, block.ID, acquisition.Owner)
	}
	in := map[int]map[string]bool{}
	queued := map[int]bool{}
	queue := make([]int, 0)
	var witness resourceExitWitness
	found := false

	for _, edge := range graph.Edges {
		if edge.From != block.ID || edge.Class != vbacfg.EdgeNormal {
			continue
		}
		if kind := applicationStateExitKind(*graph, edge.To); kind != "" {
			if resourceTransfersAtExit(proc, acquisition.Kind, kind, initialAliases) {
				continue
			}
			witness = chooseResourceWitness(witness, resourceExitWitness{Kind: kind, Line: edge.Range.StartLine})
			found = true
			continue
		}
		resourceMergeAliases(in, queued, &queue, int(edge.To), initialAliases)
	}

	for len(queue) > 0 {
		blockID := queue[0]
		queue = queue[1:]
		queued[blockID] = false
		aliases := in[blockID]
		current, ok := resourceGraphBlock(*graph, blockID)
		if !ok {
			continue
		}
		if current.Statement != nil && !current.Statement.Recovered {
			if resourceStatementReleases(acquisition.Kind, current.Statement.Text, aliases) {
				continue
			}
		}
		for _, edge := range graph.Edges {
			if int(edge.From) != blockID {
				continue
			}
			out := cloneResourceAliases(aliases)
			if current.Statement != nil && !current.Statement.Recovered && edge.Class == vbacfg.EdgeNormal {
				resourceApplyAliasAssignment(proc, acquisition.Kind, current.Statement.Text, out)
			}
			if kind := applicationStateExitKind(*graph, edge.To); kind != "" {
				if resourceTransfersAtExit(proc, acquisition.Kind, kind, out) {
					continue
				}
				witness = chooseResourceWitness(witness, resourceExitWitness{Kind: kind, Line: edge.Range.StartLine})
				found = true
				continue
			}
			resourceMergeAliases(in, queued, &queue, int(edge.To), out)
		}
	}
	return witness, found
}

// resourceFileAliasesBefore finds local numeric aliases that are definitely
// equal to the handle operand immediately before its Open statement. The
// handle itself is always present; a reassignment clears prior aliases.
func resourceFileAliasesBefore(proc sourceProcedure, acquisitionBlock vbacfg.BlockID, handle string) map[string]bool {
	graph := proc.Graph
	in := map[vbacfg.BlockID]map[string]bool{graph.Entry: {handle: true}}
	queued := map[vbacfg.BlockID]bool{graph.Entry: true}
	queue := []vbacfg.BlockID{graph.Entry}
	for len(queue) > 0 {
		blockID := queue[0]
		queue = queue[1:]
		queued[blockID] = false
		current, ok := resourceGraphBlock(*graph, int(blockID))
		if !ok || current.ID == acquisitionBlock {
			continue
		}
		aliases := in[blockID]
		for _, edge := range graph.Edges {
			if edge.From != blockID {
				continue
			}
			out := cloneResourceAliases(aliases)
			if current.Statement != nil && !current.Statement.Recovered && edge.Class == vbacfg.EdgeNormal {
				resourceApplyFileEquivalence(proc, current.Statement.Text, handle, out)
			}
			resourceMergeAliasesByBlock(in, queued, &queue, edge.To, out)
		}
	}
	if aliases, ok := in[acquisitionBlock]; ok {
		return aliases
	}
	return map[string]bool{handle: true}
}

func resourceApplyFileEquivalence(proc sourceProcedure, text, handle string, aliases map[string]bool) {
	if strings.HasPrefix(strings.ToLower(strings.TrimSpace(text)), "set ") {
		return
	}
	match := resourceValueAliasRe.FindStringSubmatch(text)
	if len(match) == 3 {
		target, source := resourceName(match[1]), resourceName(match[2])
		if !resourceLocalVariable(proc, target) {
			return
		}
		if target == handle && !aliases[source] {
			for alias := range aliases {
				delete(aliases, alias)
			}
			aliases[handle] = true
			return
		}
		if aliases[source] {
			aliases[target] = true
			return
		}
		delete(aliases, target)
		return
	}
	targetMatch := resourceValueTargetRe.FindStringSubmatch(text)
	if len(targetMatch) != 2 {
		return
	}
	target := resourceName(targetMatch[1])
	if !resourceLocalVariable(proc, target) {
		return
	}
	if target == handle {
		for alias := range aliases {
			delete(aliases, alias)
		}
		aliases[handle] = true
		return
	}
	delete(aliases, target)
}

func resourceMergeAliasesByBlock(in map[vbacfg.BlockID]map[string]bool, queued map[vbacfg.BlockID]bool, queue *[]vbacfg.BlockID, blockID vbacfg.BlockID, incoming map[string]bool) {
	current, exists := in[blockID]
	if !exists {
		in[blockID] = cloneResourceAliases(incoming)
		if !queued[blockID] {
			queued[blockID] = true
			*queue = append(*queue, blockID)
		}
		return
	}
	next := intersectResourceAliases(current, incoming)
	if resourceAliasesEqual(current, next) {
		return
	}
	in[blockID] = next
	if !queued[blockID] {
		queued[blockID] = true
		*queue = append(*queue, blockID)
	}
}

func resourceGraphBlock(graph vbacfg.Graph, id int) (vbacfg.Block, bool) {
	if id <= 0 || id > len(graph.Blocks) {
		return vbacfg.Block{}, false
	}
	return graph.Blocks[id-1], true
}

func resourceWorkbookOpenTarget(text string) (string, bool) {
	match := resourceWorkbookOpenRe.FindStringSubmatch(text)
	if len(match) != 2 {
		return "", false
	}
	return resourceName(match[1]), true
}

func resourceFileOpenHandle(text string) (string, bool) {
	match := resourceFileOpenRe.FindStringSubmatch(text)
	if len(match) != 2 {
		return "", false
	}
	return resourceName(match[1]), true
}

func resourceName(value string) string {
	return strings.ToLower(cleanIdentifier(strings.TrimSpace(value)))
}

func resourceLocalVariable(proc sourceProcedure, name string) bool {
	for declaration := range proc.Declarations.All() {
		if declaration.Scope == procedureir.ScopeLocal && strings.EqualFold(declaration.Name, name) {
			return true
		}
	}
	return false
}

func resourceWorkbookOwner(proc sourceProcedure, name string) bool {
	for declaration := range proc.Declarations.All() {
		if declaration.Scope != procedureir.ScopeLocal || !strings.EqualFold(declaration.Name, name) {
			continue
		}
		typeName := strings.TrimSpace(declaration.Type)
		if separator := strings.LastIndex(typeName, "."); separator >= 0 {
			typeName = typeName[separator+1:]
		}
		return strings.EqualFold(cleanIdentifier(strings.TrimSpace(typeName)), "Workbook")
	}
	return false
}

func resourceFileHandleOwner(proc sourceProcedure, handle string) bool {
	for _, char := range handle {
		if char < '0' || char > '9' {
			return resourceLocalVariable(proc, handle)
		}
	}
	return handle != ""
}

func resourceStatementReleases(kind resourceKind, text string, aliases map[string]bool) bool {
	if kind == resourceWorkbook {
		match := resourceCloseMethodRe.FindStringSubmatch(text)
		return len(match) == 2 && aliases[resourceName(match[1])]
	}
	match := resourceCloseRe.FindStringSubmatch(text)
	if len(match) == 0 {
		return false
	}
	arguments := strings.TrimSpace(match[1])
	if arguments == "" {
		return true
	}
	for _, part := range strings.Split(arguments, ",") {
		handle := resourceHandleRe.FindStringSubmatch(part)
		if len(handle) == 2 && aliases[resourceName(handle[1])] {
			return true
		}
	}
	return false
}

func resourceTransfersAtExit(proc sourceProcedure, kind resourceKind, exitKind string, aliases map[string]bool) bool {
	return kind == resourceWorkbook && exitKind == "normal exit" && proc.Kind == "Function" &&
		isObjectType(proc.ReturnType) && aliases[resourceName(proc.Name)]
}

func resourceApplyAliasAssignment(proc sourceProcedure, kind resourceKind, text string, aliases map[string]bool) {
	pattern := resourceSetAliasRe
	targetPattern := resourceSetTargetRe
	if kind == resourceFileHandle {
		if strings.HasPrefix(strings.ToLower(strings.TrimSpace(text)), "set ") {
			return
		}
		pattern = resourceValueAliasRe
		targetPattern = resourceValueTargetRe
	}
	match := pattern.FindStringSubmatch(text)
	if len(match) == 3 {
		target, source := resourceName(match[1]), resourceName(match[2])
		if !resourceLocalVariable(proc, target) {
			return
		}
		if aliases[source] {
			aliases[target] = true
			return
		}
		delete(aliases, target)
		return
	}
	targetMatch := targetPattern.FindStringSubmatch(text)
	if len(targetMatch) != 2 {
		return
	}
	target := resourceName(targetMatch[1])
	if resourceLocalVariable(proc, target) {
		delete(aliases, target)
	}
}

func cloneResourceAliases(in map[string]bool) map[string]bool {
	out := make(map[string]bool, len(in))
	for name := range in {
		out[name] = true
	}
	return out
}

func resourceMergeAliases(in map[int]map[string]bool, queued map[int]bool, queue *[]int, blockID int, incoming map[string]bool) {
	current, exists := in[blockID]
	if !exists {
		in[blockID] = cloneResourceAliases(incoming)
		if !queued[blockID] {
			queued[blockID] = true
			*queue = append(*queue, blockID)
		}
		return
	}
	next := intersectResourceAliases(current, incoming)
	if resourceAliasesEqual(current, next) {
		return
	}
	in[blockID] = next
	if !queued[blockID] {
		queued[blockID] = true
		*queue = append(*queue, blockID)
	}
}

func intersectResourceAliases(current, incoming map[string]bool) map[string]bool {
	next := map[string]bool{}
	for alias := range current {
		if incoming[alias] {
			next[alias] = true
		}
	}
	return next
}

func resourceAliasesEqual(left, right map[string]bool) bool {
	if len(left) != len(right) {
		return false
	}
	for alias := range left {
		if !right[alias] {
			return false
		}
	}
	return true
}

func chooseResourceWitness(current, candidate resourceExitWitness) resourceExitWitness {
	if current.Kind == "" || resourceWitnessRank(candidate.Kind) > resourceWitnessRank(current.Kind) {
		return candidate
	}
	return current
}

func resourceWitnessRank(kind string) int {
	switch kind {
	case "error exit":
		return 3
	case "unknown exit":
		return 2
	case "termination exit":
		return 1
	default:
		return 0
	}
}
