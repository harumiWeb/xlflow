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
	for _, statement := range proc.Statements {
		if statement.Recovered {
			continue
		}
		if target, ok := resourceWorkbookOpenTarget(statement.Text); ok && resourceLocalVariable(proc, target) {
			// A Function return slot is the approved v1 ownership-transfer form.
			if proc.Kind == "Function" && isObjectType(proc.ReturnType) && strings.EqualFold(target, proc.Name) {
				continue
			}
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
			witness = chooseResourceWitness(witness, resourceExitWitness{Kind: kind, Line: edge.Range.StartLine})
			found = true
			continue
		}
		resourceMergeAliases(in, queued, &queue, int(edge.To), map[string]bool{acquisition.Owner: true})
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
		if current.Statement != nil && resourceStatementReleases(acquisition.Kind, current.Statement.Text, aliases) {
			continue
		}
		if current.Statement != nil && resourceStatementTransfers(proc, acquisition.Kind, current.Statement.Text, aliases) {
			continue
		}
		for _, edge := range graph.Edges {
			if int(edge.From) != blockID {
				continue
			}
			out := cloneResourceAliases(aliases)
			if current.Statement != nil {
				// Direct local alias copies are treated as non-failing bookkeeping.
				// Acquisition remains normal-edge-only because Open itself can fail.
				resourceApplyAliasAssignment(proc, acquisition.Kind, current.Statement.Text, out)
			}
			if kind := applicationStateExitKind(*graph, edge.To); kind != "" {
				witness = chooseResourceWitness(witness, resourceExitWitness{Kind: kind, Line: edge.Range.StartLine})
				found = true
				continue
			}
			resourceMergeAliases(in, queued, &queue, int(edge.To), out)
		}
	}
	return witness, found
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
	for _, declaration := range proc.Declarations {
		if declaration.Scope == procedureir.ScopeLocal && strings.EqualFold(declaration.Name, name) {
			return true
		}
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

func resourceStatementTransfers(proc sourceProcedure, kind resourceKind, text string, aliases map[string]bool) bool {
	if kind != resourceWorkbook || proc.Kind != "Function" || !isObjectType(proc.ReturnType) {
		return false
	}
	match := resourceSetAliasRe.FindStringSubmatch(text)
	return len(match) == 3 && strings.EqualFold(resourceName(match[1]), resourceName(proc.Name)) && aliases[resourceName(match[2])]
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
	next := map[string]bool{}
	for alias := range current {
		if incoming[alias] {
			next[alias] = true
		}
	}
	if resourceAliasesEqual(current, next) {
		return
	}
	in[blockID] = next
	if !queued[blockID] {
		queued[blockID] = true
		*queue = append(*queue, blockID)
	}
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
