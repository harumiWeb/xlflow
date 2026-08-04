package cfg

import (
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func RebaseGraph(in Graph, oldBase, newBase vbaast.Range) Graph {
	out := Clone(in)
	rebased := procedureir.RebaseProcedure(procedureir.ProcedureIR{Symbol: out.Procedure}, oldBase, newBase)
	out.Procedure = rebased.Symbol
	for i := range out.Blocks {
		out.Blocks[i].Range = procedureir.RebaseRange(out.Blocks[i].Range, oldBase, newBase)
		if out.Blocks[i].Statement != nil {
			procedure := procedureir.RebaseProcedure(procedureir.ProcedureIR{Statements: []procedureir.Statement{*out.Blocks[i].Statement}}, oldBase, newBase)
			out.Blocks[i].Statement = &procedure.Statements[0]
		}
	}
	for i := range out.Edges {
		out.Edges[i].Range = procedureir.RebaseRange(out.Edges[i].Range, oldBase, newBase)
	}
	return out
}
