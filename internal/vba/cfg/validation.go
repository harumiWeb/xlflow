package cfg

import (
	"sort"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// ValidationDiagnostic is the shared, protocol-neutral projection of a
// validation fact. Adapters convert its byte range to their own coordinate
// system, while the CFG remains independent of LSP or lint packages.
type ValidationDiagnostic struct {
	Code     string
	Severity string
	Message  string
	Range    vbaast.Range
}

// ValidationDiagnostics projects certain CFG validation facts to the four
// public control-flow diagnostics. It intentionally ignores uncertain facts;
// parser recovery and conditional-compilation paths remain syntax/unknown
// flow diagnostics instead of speculative semantic errors.
func ValidationDiagnostics(document Document) []ValidationDiagnostic {
	var out []ValidationDiagnostic
	for _, graph := range document.Graphs {
		for _, fact := range graph.ValidationFacts {
			if !fact.Certain {
				continue
			}
			diagnostic, ok := validationDiagnostic(fact)
			if ok {
				out = append(out, diagnostic)
			}
		}
	}
	sort.SliceStable(out, func(i, j int) bool {
		left, right := out[i], out[j]
		if left.Range.StartByte != right.Range.StartByte {
			return left.Range.StartByte < right.Range.StartByte
		}
		if left.Range.EndByte != right.Range.EndByte {
			return left.Range.EndByte < right.Range.EndByte
		}
		return left.Code < right.Code
	})
	return out
}

func validationDiagnostic(fact ValidationFact) (ValidationDiagnostic, bool) {
	switch fact.Kind {
	case ValidationDuplicateLabel:
		return ValidationDiagnostic{
			Code: "VB055", Severity: "error", Range: fact.Range,
			Message: `Procedure label "` + fact.Name + `" is defined more than once.`,
		}, true
	case ValidationUndefinedLabel:
		return ValidationDiagnostic{
			Code: "VB056", Severity: "error", Range: fact.Range,
			Message: `Procedure label "` + fact.Name + `" is not defined.`,
		}, true
	case ValidationNextMismatch:
		message := `Next variable "` + fact.Actual + `" does not match active For variable "` + fact.Expected + `".`
		if fact.Expected == "<no active For>" {
			message = `Next variable "` + fact.Actual + `" has no compatible active For variable.`
		}
		return ValidationDiagnostic{
			Code: "VB057", Severity: "error", Range: fact.Range,
			Message: message,
		}, true
	case ValidationInvalidExit:
		return ValidationDiagnostic{
			Code: "VB058", Severity: "error", Range: fact.Range,
			Message: `Exit statement is not valid in this procedure or loop context.`,
		}, true
	default:
		return ValidationDiagnostic{}, false
	}
}

// addValidationFacts derives legality facts from the same normalized
// statements and nesting used by CFG construction. It deliberately records
// only facts that are certain from the current procedure.
func (b *builder) addValidationFacts() {
	if b.procedure.Symbol.Recovered || len(b.procedure.Symbol.ConditionalBranches) > 0 {
		return
	}
	for _, statement := range b.procedure.Statements {
		if statement.Recovered || len(statement.ConditionalBranches) > 0 {
			continue
		}
		if statement.Kind != procedureir.StatementLabel {
			continue
		}
		name := normalizedTarget(statement.Label)
		if name == "" || len(b.labels[name]) < 2 {
			continue
		}
		candidates := b.labels[name]
		if len(candidates) == 0 || candidates[0] == b.blockByID[statement.ID] || !b.certainLabelCandidates(candidates) {
			continue
		}
		rangeAt := statement.LabelRange
		if rangeAt == (vbaast.Range{}) {
			rangeAt = statement.Range
		}
		b.graph.ValidationFacts = append(b.graph.ValidationFacts, ValidationFact{
			Kind: ValidationDuplicateLabel, StatementID: statement.ID, Range: rangeAt,
			Name: statement.Label, Certain: true,
		})
	}
	for _, statement := range b.procedure.Statements {
		if statement.Recovered || len(statement.ConditionalBranches) > 0 || statement.Control == nil {
			continue
		}
		switch statement.Control.Transfer {
		case procedureir.TransferGoto, procedureir.TransferOnErrorGoto, procedureir.TransferResumeLabel:
			name := normalizedTarget(statement.Control.Target)
			if name == "" || len(b.labels[name]) != 0 {
				continue
			}
			rangeAt := statement.Control.TargetRange
			if rangeAt == (vbaast.Range{}) {
				rangeAt = statement.Range
			}
			b.graph.ValidationFacts = append(b.graph.ValidationFacts, ValidationFact{
				Kind: ValidationUndefinedLabel, StatementID: statement.ID, Range: rangeAt,
				Name: statement.Control.Target, Certain: true,
			})
		}
	}
	children := make(map[int][]int, len(b.children))
	for parent, ids := range b.children {
		children[parent] = append([]int(nil), ids...)
		sort.SliceStable(children[parent], func(i, j int) bool {
			left, right := b.stmtByID[children[parent][i]], b.stmtByID[children[parent][j]]
			if left.Range.StartByte != right.Range.StartByte {
				return left.Range.StartByte < right.Range.StartByte
			}
			return left.ID < right.ID
		})
	}
	// A parser-recovered nested For can leave its outer `Next <variable>` as
	// an ordinary call statement (the current IR has no standalone Next node).
	// In that shape the parent tree is not a reliable structured-block stack;
	// fail open for loop/Exit facts and leave the existing syntax/unknown-flow
	// diagnostics responsible for the source. Label facts above remain valid.
	structureUncertain := false
	for _, statement := range b.procedure.Statements {
		if unparsedNextStatement(statement) {
			structureUncertain = true
			break
		}
	}
	b.walkValidationChildren(children, children[0], nil, structureUncertain)
}

func unparsedNextStatement(statement procedureir.Statement) bool {
	if statement.Kind != procedureir.StatementCall && statement.Kind != procedureir.StatementUnknown {
		return false
	}
	fields := strings.Fields(strings.TrimSpace(statement.Text))
	return len(fields) > 0 && strings.EqualFold(fields[0], "next")
}

func (b *builder) certainLabelCandidates(candidates []BlockID) bool {
	for _, id := range candidates {
		statement := b.graph.block(id).Statement
		if statement == nil || statement.Recovered || len(statement.ConditionalBranches) > 0 {
			return false
		}
	}
	return true
}

type validationLoop struct {
	kind     procedureir.StatementKind
	variable string
	certain  bool
}

func (b *builder) walkValidationChildren(children map[int][]int, ids []int, loops []validationLoop, uncertain bool) {
	for _, id := range ids {
		statement := b.stmtByID[id]
		statementUncertain := uncertain || statement.Recovered || len(statement.ConditionalBranches) > 0
		currentLoops := loops
		if isValidationLoop(statement.Kind) {
			frame := validationLoop{kind: statement.Kind, certain: !statementUncertain}
			if (statement.Kind == procedureir.StatementFor || statement.Kind == procedureir.StatementForEach) && statement.Control != nil {
				frame.variable = canonicalVariable(statement.Control.LoopVariable)
				if frame.variable == "" {
					frame.certain = false
				}
			}
			currentLoops = append(append([]validationLoop(nil), loops...), frame)
			if (statement.Kind == procedureir.StatementFor || statement.Kind == procedureir.StatementForEach) &&
				!statementUncertain && frame.certain && statement.Control != nil && len(statement.Control.NextVariables) > 0 {
				b.validateNext(statement, currentLoops)
			}
		}
		if !statementUncertain && statement.Control != nil {
			if invalid := invalidExit(statement.Control.Transfer, currentLoops, b.procedure.Symbol.Kind); invalid {
				rangeAt := statement.Control.Range
				if rangeAt == (vbaast.Range{}) {
					rangeAt = statement.Range
				}
				b.graph.ValidationFacts = append(b.graph.ValidationFacts, ValidationFact{
					Kind: ValidationInvalidExit, StatementID: statement.ID, Range: rangeAt,
					Actual: string(statement.Control.Transfer), Certain: true,
				})
			}
		}
		b.walkValidationChildren(children, children[id], currentLoops, statementUncertain)
	}
}

func isValidationLoop(kind procedureir.StatementKind) bool {
	switch kind {
	case procedureir.StatementFor, procedureir.StatementForEach, procedureir.StatementDo, procedureir.StatementWhile:
		return true
	default:
		return false
	}
}

func (b *builder) validateNext(statement procedureir.Statement, loops []validationLoop) {
	control := statement.Control
	if control == nil || len(control.NextVariables) == 0 {
		return
	}
	forLoops := make([]validationLoop, 0, len(loops))
	for _, loop := range loops {
		if loop.kind == procedureir.StatementFor || loop.kind == procedureir.StatementForEach {
			forLoops = append(forLoops, loop)
		}
	}
	for index, actual := range control.NextVariables {
		if index >= len(forLoops) {
			rangeAt := statement.Range
			if index < len(control.NextVariableRanges) && control.NextVariableRanges[index] != (vbaast.Range{}) {
				rangeAt = control.NextVariableRanges[index]
			}
			b.graph.ValidationFacts = append(b.graph.ValidationFacts, ValidationFact{
				Kind: ValidationNextMismatch, StatementID: statement.ID, Range: rangeAt,
				Expected: "<no active For>", Actual: canonicalVariable(actual), Certain: true,
			})
			return
		}
		if !forLoops[len(forLoops)-1-index].certain || forLoops[len(forLoops)-1-index].variable == "" {
			return
		}
		expected := forLoops[len(forLoops)-1-index].variable
		if canonicalVariable(actual) == expected {
			continue
		}
		rangeAt := statement.Range
		if index < len(control.NextVariableRanges) && control.NextVariableRanges[index] != (vbaast.Range{}) {
			rangeAt = control.NextVariableRanges[index]
		}
		b.graph.ValidationFacts = append(b.graph.ValidationFacts, ValidationFact{
			Kind: ValidationNextMismatch, StatementID: statement.ID, Range: rangeAt,
			Expected: expected, Actual: canonicalVariable(actual), Certain: true,
		})
		return
	}
}

func invalidExit(transfer procedureir.TransferKind, loops []validationLoop, kind procedureir.ProcedureKind) bool {
	switch transfer {
	case procedureir.TransferExitFor:
		for i := len(loops) - 1; i >= 0; i-- {
			if !loops[i].certain {
				return false
			}
			if loops[i].kind == procedureir.StatementFor || loops[i].kind == procedureir.StatementForEach {
				return false
			}
		}
		return true
	case procedureir.TransferExitDo:
		for i := len(loops) - 1; i >= 0; i-- {
			if !loops[i].certain {
				return false
			}
			if loops[i].kind == procedureir.StatementDo {
				return false
			}
		}
		return true
	case procedureir.TransferExitSub:
		return kind != procedureir.ProcedureSub
	case procedureir.TransferExitFunction:
		return kind != procedureir.ProcedureFunction
	case procedureir.TransferExitProperty:
		return kind != procedureir.ProcedureProperty && kind != procedureir.ProcedurePropertyGet &&
			kind != procedureir.ProcedurePropertyLet && kind != procedureir.ProcedurePropertySet
	default:
		return false
	}
}

func canonicalVariable(value string) string {
	value = strings.TrimSpace(strings.Trim(value, "[]"))
	return strings.ToLower(strings.TrimRight(value, "$%&#@^!"))
}
