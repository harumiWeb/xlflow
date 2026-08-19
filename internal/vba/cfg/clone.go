package cfg

import (
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// CloneDocument returns a deep copy suitable for publication across cache
// boundaries.
func CloneDocument(in Document) Document {
	out := in
	out.Graphs = make([]Graph, len(in.Graphs))
	for i := range in.Graphs {
		out.Graphs[i] = Clone(in.Graphs[i])
	}
	return out
}

// Clone returns a deep copy of a graph.
func Clone(in Graph) Graph {
	out := in
	out.Procedure.Parameters = cloneParameters(in.Procedure.Parameters)
	out.Procedure.ArrayBounds = cloneArrayBounds(in.Procedure.ArrayBounds)
	out.Procedure.ConditionalBranches = append([]procedureir.ConditionalBranch(nil), in.Procedure.ConditionalBranches...)
	out.Blocks = make([]Block, len(in.Blocks))
	for i := range in.Blocks {
		out.Blocks[i] = in.Blocks[i]
		out.Blocks[i].Assignments = append([]Variable(nil), in.Blocks[i].Assignments...)
		if in.Blocks[i].Statement != nil {
			statement := *in.Blocks[i].Statement
			statement.ExpressionIDs = append([]int(nil), statement.ExpressionIDs...)
			statement.Target = cloneExpression(statement.Target)
			statement.Value = cloneExpression(statement.Value)
			statement.Condition = cloneExpression(statement.Condition)
			statement.ConditionalBranches = append([]procedureir.ConditionalBranch(nil), statement.ConditionalBranches...)
			if statement.Control != nil {
				control := *statement.Control
				control.NextVariables = append([]string(nil), statement.Control.NextVariables...)
				control.NextVariableRanges = append([]vbaast.Range(nil), statement.Control.NextVariableRanges...)
				statement.Control = &control
			}
			out.Blocks[i].Statement = &statement
		}
	}
	out.Edges = append([]Edge(nil), in.Edges...)
	out.UnknownFlowSources = append([]BlockID(nil), in.UnknownFlowSources...)
	out.ValidationFacts = append([]ValidationFact(nil), in.ValidationFacts...)
	out.query = buildQueryIndex(out)
	return out
}

func cloneParameters(in []procedureir.Parameter) []procedureir.Parameter {
	out := append([]procedureir.Parameter(nil), in...)
	for i := range out {
		out[i].DefaultRange = cloneRangePointer(in[i].DefaultRange)
		out[i].BoundsRange = cloneRangePointer(in[i].BoundsRange)
		out[i].ArrayBounds = cloneArrayBounds(in[i].ArrayBounds)
	}
	return out
}

func cloneArrayBounds(in []procedureir.ArrayBound) []procedureir.ArrayBound {
	out := append([]procedureir.ArrayBound(nil), in...)
	for i := range out {
		out[i].LowerRange = cloneRangePointer(in[i].LowerRange)
		out[i].UpperRange = cloneRangePointer(in[i].UpperRange)
	}
	return out
}

func cloneRangePointer(in *vbaast.Range) *vbaast.Range {
	if in == nil {
		return nil
	}
	out := *in
	return &out
}

func cloneExpression(in *procedureir.Expression) *procedureir.Expression {
	if in == nil {
		return nil
	}
	out := *in
	out.Children = append([]int(nil), in.Children...)
	return &out
}
