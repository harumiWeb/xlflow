package procedureir

func Clone(in DocumentIR) DocumentIR {
	out := in
	out.ModuleAttributes = append([]ModuleAttribute(nil), in.ModuleAttributes...)
	out.Declarations = append([]Declaration(nil), in.Declarations...)
	for i := range out.Declarations {
		out.Declarations[i].Parameters = append([]Parameter(nil), in.Declarations[i].Parameters...)
	}
	out.TypeReferences = make([]TypeReference, len(in.TypeReferences))
	for i := range in.TypeReferences {
		out.TypeReferences[i] = in.TypeReferences[i]
		if in.TypeReferences[i].Caller != nil {
			caller := *in.TypeReferences[i].Caller
			out.TypeReferences[i].Caller = &caller
		}
	}
	out.Procedures = make([]ProcedureIR, len(in.Procedures))
	for i := range in.Procedures {
		out.Procedures[i] = cloneProcedure(in.Procedures[i])
	}
	return out
}

// CloneDocumentIR is the explicit-name form of Clone.
func CloneDocumentIR(in DocumentIR) DocumentIR {
	return Clone(in)
}

// CloneProcedureIR returns a deep copy of one procedure.
func CloneProcedureIR(in ProcedureIR) ProcedureIR {
	return cloneProcedure(in)
}

// CloneCallSite returns a deep copy of one call site.
func CloneCallSite(in CallSite) CallSite {
	return cloneCall(in)
}

func cloneProcedure(in ProcedureIR) ProcedureIR {
	out := in
	out.Symbol.Parameters = append([]Parameter(nil), in.Symbol.Parameters...)
	out.Declarations = append([]Declaration(nil), in.Declarations...)
	for i := range out.Declarations {
		out.Declarations[i].Parameters = append([]Parameter(nil), in.Declarations[i].Parameters...)
	}
	out.Statements = make([]Statement, len(in.Statements))
	for i := range in.Statements {
		out.Statements[i] = in.Statements[i]
		out.Statements[i].ExpressionIDs = append([]int(nil), in.Statements[i].ExpressionIDs...)
		out.Statements[i].Target = cloneExpressionPointer(in.Statements[i].Target)
		out.Statements[i].Value = cloneExpressionPointer(in.Statements[i].Value)
		out.Statements[i].Condition = cloneExpressionPointer(in.Statements[i].Condition)
		if in.Statements[i].Control != nil {
			control := *in.Statements[i].Control
			out.Statements[i].Control = &control
		}
	}
	out.Expressions = make([]Expression, len(in.Expressions))
	for i := range in.Expressions {
		out.Expressions[i] = in.Expressions[i]
		out.Expressions[i].Children = append([]int(nil), in.Expressions[i].Children...)
	}
	out.Calls = make([]CallSite, len(in.Calls))
	for i := range in.Calls {
		out.Calls[i] = cloneCall(in.Calls[i])
	}
	out.Accesses = make([]VariableAccess, len(in.Accesses))
	for i := range in.Accesses {
		out.Accesses[i] = in.Accesses[i]
		out.Accesses[i].Resolution.Candidates = append([]Candidate(nil), in.Accesses[i].Resolution.Candidates...)
	}
	return out
}

func cloneExpressionPointer(in *Expression) *Expression {
	if in == nil {
		return nil
	}
	out := *in
	out.Children = append([]int(nil), in.Children...)
	return &out
}

func cloneCall(in CallSite) CallSite {
	out := in
	if in.Callee.Receiver != nil {
		receiver := *in.Callee.Receiver
		out.Callee.Receiver = &receiver
	}
	out.Arguments.Named = append([]NamedArgument(nil), in.Arguments.Named...)
	out.Arguments.ExpressionIDs = append([]int(nil), in.Arguments.ExpressionIDs...)
	out.Resolution.Candidates = append([]Candidate(nil), in.Resolution.Candidates...)
	return out
}
