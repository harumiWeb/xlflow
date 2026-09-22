package analyze

import (
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// deadStoreCandidate identifies a scalar local assignment whose value is not
// observed on any path on which that assignment completes.  The candidate is
// intentionally a small, analyzer-local value: callers decide how to project
// it into batch, realtime, or LSP diagnostics.
type deadStoreCandidate struct {
	StatementID int
	Name        string
	Range       vbaast.Range
}

func (a Analyzer) deadStoreFindings(file parsedFile, proc sourceProcedure) []Finding {
	findings := make([]Finding, 0)
	for _, candidate := range deadStoreCandidates(proc) {
		line := candidate.Range.StartLine
		if line < 1 {
			continue
		}
		finding := a.simpleFinding(file, proc, line, "VBA256", "warning",
			"Assignment to "+candidate.Name+" is never read before the value is overwritten or the procedure exits.",
			"The assigned scalar value is not observed on any completed control-flow path.",
			"Remove the assignment or use the value before assigning it again.")
		finding.Column = candidate.Range.StartColumn + 1
		finding.EndLine = candidate.Range.EndLine
		finding.EndColumn = candidate.Range.EndColumn + 1
		findings = append(findings, finding)
	}
	return findings
}

// deadStoreCandidates performs a procedure-local backward liveness analysis.
// It is deliberately limited to explicit scalar local variables and plain
// StatementAssignment writes.  All CFG edges are considered; exceptional
// edges preserve the block input because an assignment may not have completed
// before the exception was raised.
func deadStoreCandidates(proc sourceProcedure) []deadStoreCandidate {
	if proc.Graph == nil || proc.Graph.Entry == 0 {
		return nil
	}

	eligible := deadStoreEligibleDeclarations(proc.Declarations)
	if len(eligible) == 0 {
		return nil
	}

	view := proc.Graph.View(vbacfg.EdgeFilter{})
	reachable := make(map[vbacfg.BlockID]bool)
	for _, id := range view.Reachable() {
		reachable[id] = true
	}
	if len(reachable) == 0 {
		return nil
	}

	// An unknown-flow source can transfer control to an unmodeled statement.
	// There is no sound finite successor set for such a source, so fail open
	// for the whole procedure rather than report a potentially false dead store.
	for _, source := range proc.Graph.UnknownFlowSources {
		if reachable[source] {
			return nil
		}
	}

	accessesByStatement := make(map[int][]procedureir.VariableAccess)
	for access := range proc.Accesses.All() {
		if access.Mode == procedureir.AccessReadWrite {
			// A ByRef/read-write operation can observe or mutate the value outside
			// this local proof. Keep the whole procedure fail-open rather than
			// treating only the named variable as isolated.
			return nil
		}
		accessesByStatement[access.StatementID] = append(accessesByStatement[access.StatementID], access)
	}

	writes := make(map[vbacfg.BlockID]deadStoreCandidate)
	readsByStatement := make(map[int]map[string]bool)
	for statementID, accesses := range accessesByStatement {
		for _, access := range accesses {
			name := deadStoreCanonicalName(access.Name)
			if name == "" || !eligible[name] || !deadStoreLocalAccess(access.Scope) {
				continue
			}
			if access.Mode != procedureir.AccessWrite {
				if readsByStatement[statementID] == nil {
					readsByStatement[statementID] = make(map[string]bool)
				}
				readsByStatement[statementID][name] = true
			}
		}
	}

	for _, block := range proc.Graph.Blocks {
		if !reachable[block.ID] || block.Kind != vbacfg.BlockStatement || block.Statement == nil {
			continue
		}
		statement := block.Statement
		if statement.Recovered || statement.Kind != procedureir.StatementAssignment ||
			statement.Target == nil || statement.Target.Recovered ||
			statement.Target.Kind != procedureir.ExpressionIdentifier {
			continue
		}
		name := deadStoreCanonicalName(statement.Target.Text)
		if name == "" || !eligible[name] {
			continue
		}
		if deadStorePlainWrite(accessesByStatement[statement.ID], name) {
			writes[block.ID] = deadStoreCandidate{
				StatementID: statement.ID,
				Name:        name,
				Range:       statement.Range,
			}
		}
	}

	liveIn := make(map[vbacfg.BlockID]map[string]bool, len(proc.Graph.Blocks))
	liveOut := make(map[vbacfg.BlockID]map[string]bool, len(proc.Graph.Blocks))
	for _, block := range proc.Graph.Blocks {
		if reachable[block.ID] {
			liveIn[block.ID] = make(map[string]bool)
			liveOut[block.ID] = make(map[string]bool)
		}
	}

	queue := make([]vbacfg.BlockID, 0, len(reachable))
	inQueue := make(map[vbacfg.BlockID]bool, len(reachable))
	blocksByID := make(map[vbacfg.BlockID]vbacfg.Block, len(proc.Graph.Blocks))
	for _, block := range proc.Graph.Blocks {
		blocksByID[block.ID] = block
	}
	for id := range reachable {
		queue = append(queue, id)
		inQueue[id] = true
	}
	for len(queue) > 0 {
		blockID := queue[len(queue)-1]
		queue = queue[:len(queue)-1]
		inQueue[blockID] = false
		block, ok := blocksByID[blockID]
		if !ok {
			continue
		}

		normalAfter := make(map[string]bool)
		exceptionalAfter := make(map[string]bool)
		view.ForEachOutgoing(blockID, func(edge vbacfg.Edge) bool {
			target := liveIn[edge.To]
			if edge.Class == vbacfg.EdgeExceptional {
				deadStoreAddSet(exceptionalAfter, target)
			} else {
				deadStoreAddSet(normalAfter, target)
			}
			return true
		})

		out := make(map[string]bool, len(normalAfter)+len(exceptionalAfter))
		deadStoreAddSet(out, normalAfter)
		deadStoreAddSet(out, exceptionalAfter)

		in := make(map[string]bool, len(out))
		// A write is complete only on normal edges.  Exceptional successors
		// use the pre-statement input, so their liveness is added below after
		// applying normal defs.
		deadStoreAddSet(in, normalAfter)
		if write, ok := writes[block.ID]; ok {
			delete(in, write.Name)
		}
		deadStoreAddSet(in, exceptionalAfter)
		if block.Statement != nil {
			deadStoreAddSet(in, readsByStatement[block.Statement.ID])
		}

		outChanged := !deadStoreSameSet(liveOut[blockID], out)
		inChanged := !deadStoreSameSet(liveIn[blockID], in)
		if outChanged {
			liveOut[blockID] = out
		}
		if inChanged {
			liveIn[blockID] = in
		}
		if outChanged || inChanged {
			view.ForEachIncoming(blockID, func(edge vbacfg.Edge) bool {
				if reachable[edge.From] && !inQueue[edge.From] {
					queue = append(queue, edge.From)
					inQueue[edge.From] = true
				}
				return true
			})
		}
	}

	candidates := make([]deadStoreCandidate, 0, len(writes))
	for _, block := range proc.Graph.Blocks {
		write, ok := writes[block.ID]
		if !ok {
			continue
		}
		after := make(map[string]bool)
		view.ForEachOutgoing(block.ID, func(edge vbacfg.Edge) bool {
			if edge.Class == vbacfg.EdgeExceptional {
				return true
			}
			if edge.Uncertain || edge.Kind == vbacfg.EdgeUnknown {
				after[write.Name] = true
				return false
			}
			target := liveIn[edge.To]
			if target[write.Name] {
				after[write.Name] = true
				return false
			}
			return true
		})
		if !after[write.Name] {
			candidates = append(candidates, write)
		}
	}
	return candidates
}

func deadStoreEligibleDeclarations(declarations readOnlySpan[procedureir.Declaration]) map[string]bool {
	eligible := make(map[string]bool)
	for declaration := range declarations.All() {
		if declaration.Scope != procedureir.ScopeLocal || declaration.Kind == "return_slot" || declaration.IsStatic || declaration.IsArray ||
			declaration.IsObject || declaration.IsNew || declaration.IsConst {
			continue
		}
		if strings.EqualFold(strings.TrimSpace(declaration.Type), "") ||
			strings.EqualFold(strings.TrimSpace(declaration.Type), "variant") ||
			declaration.ValueShape == procedureir.ValueShapeVariant ||
			declaration.ValueShape == procedureir.ValueShapeFixedArray ||
			declaration.ValueShape == procedureir.ValueShapeDynamicArray {
			continue
		}
		name := deadStoreCanonicalName(declaration.Name)
		if name != "" {
			eligible[name] = true
		}
	}
	return eligible
}

func deadStorePlainWrite(accesses []procedureir.VariableAccess, name string) bool {
	writes := 0
	for _, access := range accesses {
		if deadStoreCanonicalName(access.Name) != name || !deadStoreLocalAccess(access.Scope) {
			continue
		}
		switch access.Mode {
		case procedureir.AccessWrite:
			writes++
		case procedureir.AccessReadWrite:
			// A read-write target is not the strict simple assignment shape.
			return false
		}
	}
	return writes == 1
}

func deadStoreLocalAccess(scope procedureir.SymbolScope) bool {
	return scope == procedureir.ScopeLocal
}

func deadStoreCanonicalName(name string) string {
	return strings.ToLower(strings.TrimSpace(cleanIdentifier(name)))
}

func deadStoreAddSet(dst, src map[string]bool) {
	for name := range src {
		dst[name] = true
	}
}

func deadStoreSameSet(a, b map[string]bool) bool {
	if len(a) != len(b) {
		return false
	}
	for name := range a {
		if !b[name] {
			return false
		}
	}
	return true
}
