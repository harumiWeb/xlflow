package procedureir

import vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"

// RebaseProcedure returns a deep copy whose source ranges move from oldBase to
// newBase. IDs remain procedure-local and deterministic.
func RebaseProcedure(in ProcedureIR, oldBase, newBase vbaast.Range) ProcedureIR {
	out := CloneProcedureIR(in)
	rebase := func(r vbaast.Range) vbaast.Range { return RebaseRange(r, oldBase, newBase) }
	out.Symbol.DeclarationRange = rebase(out.Symbol.DeclarationRange)
	out.Symbol.BodyRange = rebase(out.Symbol.BodyRange)
	for i := range out.Symbol.ConditionalBranches {
		out.Symbol.ConditionalBranches[i].Range = rebase(out.Symbol.ConditionalBranches[i].Range)
	}
	for i := range out.Symbol.Parameters {
		out.Symbol.Parameters[i].Range = rebase(out.Symbol.Parameters[i].Range)
		out.Symbol.Parameters[i].DefaultRange = rebasePointer(out.Symbol.Parameters[i].DefaultRange, rebase)
		out.Symbol.Parameters[i].BoundsRange = rebasePointer(out.Symbol.Parameters[i].BoundsRange, rebase)
		for j := range out.Symbol.Parameters[i].ArrayBounds {
			bound := &out.Symbol.Parameters[i].ArrayBounds[j]
			bound.Range = rebase(bound.Range)
			bound.LowerRange = rebasePointer(bound.LowerRange, rebase)
			bound.UpperRange = rebasePointer(bound.UpperRange, rebase)
		}
	}
	for i := range out.Declarations {
		out.Declarations[i].Range = rebase(out.Declarations[i].Range)
		for j := range out.Declarations[i].ConditionalBranches {
			out.Declarations[i].ConditionalBranches[j].Range = rebase(out.Declarations[i].ConditionalBranches[j].Range)
		}
		for j := range out.Declarations[i].Parameters {
			parameter := &out.Declarations[i].Parameters[j]
			parameter.Range = rebase(parameter.Range)
			parameter.DefaultRange = rebasePointer(parameter.DefaultRange, rebase)
			parameter.BoundsRange = rebasePointer(parameter.BoundsRange, rebase)
			for k := range parameter.ArrayBounds {
				bound := &parameter.ArrayBounds[k]
				bound.Range = rebase(bound.Range)
				bound.LowerRange = rebasePointer(bound.LowerRange, rebase)
				bound.UpperRange = rebasePointer(bound.UpperRange, rebase)
			}
		}
	}
	for i := range out.Statements {
		out.Statements[i].Range = rebase(out.Statements[i].Range)
		out.Statements[i].LabelRange = rebase(out.Statements[i].LabelRange)
		for j := range out.Statements[i].ConditionalBranches {
			out.Statements[i].ConditionalBranches[j].Range = rebase(out.Statements[i].ConditionalBranches[j].Range)
		}
		rebaseExpression := func(expression *Expression) {
			if expression != nil {
				expression.Range = rebase(expression.Range)
			}
		}
		rebaseExpression(out.Statements[i].Target)
		rebaseExpression(out.Statements[i].Value)
		rebaseExpression(out.Statements[i].Condition)
		if out.Statements[i].Control != nil {
			out.Statements[i].Control.Range = rebase(out.Statements[i].Control.Range)
			out.Statements[i].Control.TargetRange = rebase(out.Statements[i].Control.TargetRange)
			out.Statements[i].Control.LoopVariableRange = rebase(out.Statements[i].Control.LoopVariableRange)
			for j := range out.Statements[i].Control.NextVariableRanges {
				out.Statements[i].Control.NextVariableRanges[j] = rebase(out.Statements[i].Control.NextVariableRanges[j])
			}
		}
	}
	for i := range out.Expressions {
		out.Expressions[i].Range = rebase(out.Expressions[i].Range)
	}
	for i := range out.Calls {
		out.Calls[i].Range = rebase(out.Calls[i].Range)
	}
	for i := range out.RaiseEvents {
		out.RaiseEvents[i].Range = rebase(out.RaiseEvents[i].Range)
		for j := range out.RaiseEvents[i].ConditionalBranches {
			out.RaiseEvents[i].ConditionalBranches[j].Range = rebase(out.RaiseEvents[i].ConditionalBranches[j].Range)
		}
	}
	for i := range out.Accesses {
		out.Accesses[i].Range = rebase(out.Accesses[i].Range)
	}
	return out
}

func rebasePointer(pointer *vbaast.Range, rebase func(vbaast.Range) vbaast.Range) *vbaast.Range {
	if pointer == nil {
		return nil
	}
	rng := rebase(*pointer)
	return &rng
}

func RebaseRange(r, oldBase, newBase vbaast.Range) vbaast.Range {
	if r == (vbaast.Range{}) {
		return r
	}
	lineDeltaStart := r.StartLine - oldBase.StartLine
	lineDeltaEnd := r.EndLine - oldBase.StartLine
	r.StartByte = newBase.StartByte + r.StartByte - oldBase.StartByte
	r.EndByte = newBase.StartByte + r.EndByte - oldBase.StartByte
	r.StartLine = newBase.StartLine + lineDeltaStart
	r.EndLine = newBase.StartLine + lineDeltaEnd
	if lineDeltaStart == 0 {
		r.StartColumn = newBase.StartColumn + r.StartColumn - oldBase.StartColumn
	}
	if lineDeltaEnd == 0 {
		r.EndColumn = newBase.StartColumn + r.EndColumn - oldBase.StartColumn
	}
	return r
}

func RebaseTypeReference(in TypeReference, oldBase, newBase vbaast.Range) TypeReference {
	in.Range = RebaseRange(in.Range, oldBase, newBase)
	if in.Caller != nil {
		caller := *in.Caller
		in.Caller = &caller
	}
	return in
}
