// Package cfg builds protocol-neutral, conservative control-flow graphs from
// normalized VBA procedure IR.
package cfg

import (
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

type BlockID int
type EdgeID int

type BlockKind string

const (
	BlockEntry           BlockKind = "entry"
	BlockNormalExit      BlockKind = "normal_exit"
	BlockExceptionalExit BlockKind = "exceptional_exit"
	BlockTerminationExit BlockKind = "termination_exit"
	BlockUnknownExit     BlockKind = "unknown_exit"
	BlockStatement       BlockKind = "statement"
)

type EdgeClass string

const (
	EdgeNormal      EdgeClass = "normal"
	EdgeExceptional EdgeClass = "exceptional"
)

type EdgeKind string

const (
	EdgeFallthrough   EdgeKind = "fallthrough"
	EdgeBranchTrue    EdgeKind = "branch_true"
	EdgeBranchFalse   EdgeKind = "branch_false"
	EdgeCase          EdgeKind = "case"
	EdgeLoopBody      EdgeKind = "loop_body"
	EdgeLoopExit      EdgeKind = "loop_exit"
	EdgeLoopBack      EdgeKind = "loop_back"
	EdgeGoto          EdgeKind = "goto"
	EdgeProcedureExit EdgeKind = "procedure_exit"
	EdgeTermination   EdgeKind = "termination"
	EdgeError         EdgeKind = "error"
	EdgeResume        EdgeKind = "resume"
	EdgeUnknown       EdgeKind = "unknown"
)

// ValidationKind identifies a protocol-neutral control-flow legality fact.
// Diagnostic packages map these facts to user-facing rule IDs.
type ValidationKind string

const (
	ValidationDuplicateLabel ValidationKind = "duplicate_label"
	ValidationUndefinedLabel ValidationKind = "undefined_label"
	ValidationNextMismatch   ValidationKind = "next_mismatch"
	ValidationInvalidExit    ValidationKind = "invalid_exit"
)

type ValidationFact struct {
	Kind        ValidationKind `json:"kind"`
	StatementID int            `json:"statementId"`
	Range       vbaast.Range   `json:"range"`
	Name        string         `json:"name,omitempty"`
	Expected    string         `json:"expected,omitempty"`
	Actual      string         `json:"actual,omitempty"`
	Certain     bool           `json:"certain"`
}

// Block is either one normalized VBA statement or a synthetic graph endpoint.
type Block struct {
	ID          BlockID                `json:"id"`
	Kind        BlockKind              `json:"kind"`
	StatementID int                    `json:"statementId,omitempty"`
	Statement   *procedureir.Statement `json:"statement,omitempty"`
	Range       vbaast.Range           `json:"range"`
	Assignments []Variable             `json:"assignments,omitempty"`
}

// Edge records uncertainty independently from the normal/exceptional class.
// StatementID and Range identify the source transition in the original VBA.
type Edge struct {
	ID          EdgeID       `json:"id"`
	From        BlockID      `json:"from"`
	To          BlockID      `json:"to"`
	Kind        EdgeKind     `json:"kind"`
	Class       EdgeClass    `json:"class"`
	Uncertain   bool         `json:"uncertain,omitempty"`
	StatementID int          `json:"statementId,omitempty"`
	Range       vbaast.Range `json:"range"`
}

// Variable is the case-insensitive identity used by definite-assignment
// queries. Scope remains part of the key so local and module bindings do not
// alias.
type Variable struct {
	Scope procedureir.SymbolScope `json:"scope"`
	Name  string                  `json:"name"`
}

func (v Variable) canonical() Variable {
	v.Name = strings.ToLower(v.Name)
	return v
}

// Graph owns all data for one ProcedureIR. IDs are deterministic for an
// identical input but are meaningful only within this graph revision.
type Graph struct {
	Procedure          procedureir.ProcedureSymbol `json:"procedure"`
	Blocks             []Block                     `json:"blocks"`
	Edges              []Edge                      `json:"edges"`
	UnknownFlowSources []BlockID                   `json:"unknownFlowSources,omitempty"`
	ValidationFacts    []ValidationFact            `json:"validationFacts,omitempty"`
	Entry              BlockID                     `json:"entry"`
	NormalExit         BlockID                     `json:"normalExit"`
	ExceptionalExit    BlockID                     `json:"exceptionalExit"`
	TerminationExit    BlockID                     `json:"terminationExit"`
	UnknownExit        BlockID                     `json:"unknownExit"`
	query              *queryIndex
}

type Document struct {
	Path   string  `json:"path"`
	Graphs []Graph `json:"graphs"`
}

// EdgeFilter selects graph views. Zero value is the conservative guarantee
// view: every normal and exceptional edge, including uncertain edges.
type EdgeFilter struct {
	NormalOnly bool
}

func (f EdgeFilter) accepts(edge Edge) bool {
	if f.NormalOnly && edge.Class != EdgeNormal {
		return false
	}
	return true
}

// ExitSelection controls which synthetic exits participate in path queries.
// Zero value selects every exit, including unknown analysis outcomes.
type ExitSelection struct {
	Normal      bool
	Exceptional bool
	Termination bool
	Unknown     bool
}

func (s ExitSelection) ids(g Graph) []BlockID {
	if !s.Normal && !s.Exceptional && !s.Termination && !s.Unknown {
		return []BlockID{g.NormalExit, g.ExceptionalExit, g.TerminationExit, g.UnknownExit}
	}
	var out []BlockID
	if s.Normal {
		out = append(out, g.NormalExit)
	}
	if s.Exceptional {
		out = append(out, g.ExceptionalExit)
	}
	if s.Termination {
		out = append(out, g.TerminationExit)
	}
	if s.Unknown {
		out = append(out, g.UnknownExit)
	}
	return out
}
