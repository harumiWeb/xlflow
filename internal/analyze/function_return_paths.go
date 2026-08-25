package analyze

import (
	"fmt"
	"sort"
	"strings"

	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

type returnAssignmentKind uint8

const (
	returnAssignmentNone returnAssignmentKind = iota
	returnAssignmentValid
	returnAssignmentInvalid
)

type returnPathState uint8

const (
	returnPathUnassigned returnPathState = iota + 1
	returnPathAssigned
	returnPathInvalid
)

type returnPathWitness struct {
	Kind     string
	Line     int
	Invalid  bool
	ViaError bool
}

func (a Analyzer) functionReturnPathFindings(file parsedFile, proc sourceProcedure) []Finding {
	if !a.Config.Analyze.DetectFunctionReturnPath || proc.Name == "" || proc.Graph == nil {
		return nil
	}
	if proc.ProcedureKind != procedureir.ProcedureFunction && proc.ProcedureKind != procedureir.ProcedurePropertyGet {
		return nil
	}

	graph := proc.Graph.WithoutNormalErrRaiseContinuationView()
	if !graph.IsReachable(graph.NormalExit()) {
		return nil
	}

	variable := returnSlotVariable(proc.Name)
	returnKinds := map[int]returnAssignmentKind{}
	definite := graph.DefiniteAssignmentsWith(func(block vbacfg.Block, assigned vbacfg.Variable) bool {
		if block.Statement == nil {
			return true
		}
		kind, cached := returnKinds[block.Statement.ID]
		if !cached {
			kind = returnAssignmentKindForStatement(proc, *block.Statement)
			returnKinds[block.Statement.ID] = kind
		}
		if kind != returnAssignmentInvalid {
			return true
		}
		return assigned.Scope != variable.Scope || !strings.EqualFold(assigned.Name, variable.Name)
	})
	if hasReturnVariable(definite[graph.NormalExit()], variable) {
		return nil
	}

	witnesses := returnPathWitnesses(graph, proc)
	if len(witnesses) == 0 {
		witnesses = []returnPathWitness{{
			Kind: "unknown control-flow path to normal exit",
			Line: proc.StartLine,
		}}
	}
	if isObjectType(proc.ReturnType) && allReturnWitnessesInvalid(witnesses) {
		// VBA103 is the existing, more precise batch diagnostic for this case.
		return nil
	}

	witness := selectReturnPathWitness(witnesses)
	line := proc.StartLine
	if line <= 0 {
		line = witness.Line
	}
	path := witness.Kind
	message := fmt.Sprintf("%s may exit through %s without a valid return assignment.", proc.Name, path)
	reason := fmt.Sprintf("The %s does not reach a valid assignment to the %s return slot; VBA default initialization is not an explicit return assignment.", path, proc.Name)
	if witness.Line > 0 {
		reason += fmt.Sprintf(" The representative exit is at line %d.", witness.Line)
	}
	if witness.Invalid {
		reason += " The return assignment uses the wrong VBA assignment syntax for its declared type."
	}

	suggestion := "Assign " + proc.Name + " on every successful return path."
	switch {
	case isObjectType(proc.ReturnType):
		suggestion = "Use Set " + proc.Name + " = ... on every successful return path."
	case definiteValueReturnType(proc.ReturnType):
		suggestion = "Use " + proc.Name + " = ... (or Let " + proc.Name + " = ...) on every successful return path."
	}
	return []Finding{a.simpleFinding(file, proc, line, "VBA210", "warning", message, reason, suggestion)}
}

func returnSlotVariable(name string) vbacfg.Variable {
	return vbacfg.Variable{Scope: procedureir.ScopeLocal, Name: strings.ToLower(name)}
}

func returnAssignmentKindForStatement(proc sourceProcedure, statement procedureir.Statement) returnAssignmentKind {
	if statement.Kind != procedureir.StatementAssignment && statement.Kind != procedureir.StatementSet {
		return returnAssignmentNone
	}
	if !statementWritesReturnSlot(proc, statement.ID) {
		return returnAssignmentNone
	}

	switch {
	case isObjectType(proc.ReturnType):
		if statement.Kind == procedureir.StatementSet {
			return returnAssignmentValid
		}
		return returnAssignmentInvalid
	case definiteValueReturnType(proc.ReturnType):
		if statement.Kind == procedureir.StatementAssignment {
			return returnAssignmentValid
		}
		return returnAssignmentInvalid
	default:
		return returnAssignmentValid
	}
}

func statementWritesReturnSlot(proc sourceProcedure, statementID int) bool {
	for access := range proc.Accesses.All() {
		if access.StatementID != statementID || access.Scope != procedureir.ScopeLocal {
			continue
		}
		if access.Mode != procedureir.AccessWrite && access.Mode != procedureir.AccessReadWrite {
			continue
		}
		if strings.EqualFold(access.Name, proc.Name) {
			return true
		}
	}
	return false
}

func definiteValueReturnType(typ string) bool {
	switch strings.ToLower(strings.TrimSpace(typ)) {
	case "boolean", "byte", "currency", "date", "decimal", "double", "integer", "long", "longlong", "longptr", "single", "string":
		return true
	default:
		return false
	}
}

func hasReturnVariable(assignments []vbacfg.Variable, variable vbacfg.Variable) bool {
	for _, assigned := range assignments {
		if assigned.Scope == variable.Scope && strings.EqualFold(assigned.Name, variable.Name) {
			return true
		}
	}
	return false
}

func returnPathWitnesses(graph vbacfg.CFGView, proc sourceProcedure) []returnPathWitness {
	type pathKey struct {
		block    vbacfg.BlockID
		state    returnPathState
		viaError bool
	}
	queue := []pathKey{{block: graph.Entry(), state: returnPathUnassigned}}
	seen := map[pathKey]bool{queue[0]: true}
	var witnesses []returnPathWitness

	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		block, ok := returnPathBlock(graph, current.block)
		if !ok {
			continue
		}
		graph.ForEachOutgoing(current.block, func(edge vbacfg.Edge) bool {
			nextState := current.state
			if edge.Class == vbacfg.EdgeNormal && block.Statement != nil {
				switch returnAssignmentKindForStatement(proc, *block.Statement) {
				case returnAssignmentValid:
					nextState = returnPathAssigned
				case returnAssignmentInvalid:
					nextState = returnPathInvalid
				}
			}
			viaError := current.viaError || edge.Class == vbacfg.EdgeExceptional
			if edge.To == graph.NormalExit() && nextState != returnPathAssigned {
				witnesses = append(witnesses, returnPathWitness{
					Kind:     returnExitKind(graph, edge, viaError),
					Line:     returnExitLine(graph, edge, proc),
					Invalid:  nextState == returnPathInvalid,
					ViaError: viaError,
				})
			}
			next := pathKey{block: edge.To, state: nextState, viaError: viaError}
			if seen[next] {
				return true
			}
			seen[next] = true
			queue = append(queue, next)
			return true
		})
	}
	return witnesses
}

func returnPathBlock(graph vbacfg.CFGView, id vbacfg.BlockID) (vbacfg.Block, bool) {
	return graph.BlockByID(id)
}

func returnExitKind(graph vbacfg.CFGView, edge vbacfg.Edge, viaError bool) string {
	if viaError {
		return "error-handler path to normal exit"
	}
	if edge.Kind == vbacfg.EdgeProcedureExit {
		if block, ok := returnPathBlock(graph, edge.From); ok && block.Statement != nil && block.Statement.Control != nil {
			switch block.Statement.Control.Transfer {
			case procedureir.TransferExitFunction:
				return "Exit Function"
			case procedureir.TransferExitProperty:
				return "Exit Property"
			}
		}
		return "explicit procedure exit"
	}
	if edge.Kind == vbacfg.EdgeFallthrough {
		return "normal fallthrough"
	}
	return "normal exit"
}

func returnExitLine(graph vbacfg.CFGView, edge vbacfg.Edge, proc sourceProcedure) int {
	if edge.Range.StartLine > 0 {
		return edge.Range.StartLine
	}
	if block, ok := returnPathBlock(graph, edge.From); ok && block.Range.StartLine > 0 {
		return block.Range.StartLine
	}
	return proc.EndLine
}

func allReturnWitnessesInvalid(witnesses []returnPathWitness) bool {
	if len(witnesses) == 0 {
		return false
	}
	for _, witness := range witnesses {
		if !witness.Invalid {
			return false
		}
	}
	return true
}

func selectReturnPathWitness(witnesses []returnPathWitness) returnPathWitness {
	sort.SliceStable(witnesses, func(i, j int) bool {
		rank := func(w returnPathWitness) int {
			switch {
			case w.ViaError:
				return 0
			case strings.HasPrefix(w.Kind, "Exit "):
				return 1
			case w.Kind == "normal fallthrough":
				return 2
			default:
				return 3
			}
		}
		left, right := rank(witnesses[i]), rank(witnesses[j])
		if left != right {
			return left < right
		}
		if witnesses[i].Line != witnesses[j].Line {
			return witnesses[i].Line < witnesses[j].Line
		}
		return witnesses[i].Kind < witnesses[j].Kind
	})
	return witnesses[0]
}
