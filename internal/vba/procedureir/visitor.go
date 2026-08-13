package procedureir

import (
	"sort"
	"strconv"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

type visitContext struct {
	procedure      int
	statementID    int
	parentExprID   int
	branch         BranchRole
	metadata       bool
	accessMode     AccessMode
	callIndex      int
	suppressCall   bool
	inBody         bool
	conditional    []ConditionalBranch
	enumName       string
	enumVisibility string
}

type argumentFact struct {
	rng        vbaast.Range
	name       string
	valueText  string
	valueRange vbaast.Range
}

type callFact struct {
	arguments []argumentFact
}

type singleVisitor struct {
	builder     *documentBuilder
	document    DocumentIR
	callFacts   [][]callFact
	moduleScope map[string]SymbolScope
	visited     uint64
}

func (b *documentBuilder) buildSinglePass(root *tree_sitter.Node) DocumentIR {
	v := singleVisitor{
		builder: b,
		document: DocumentIR{
			Path: b.file, ModuleName: b.moduleName, ModuleKind: b.moduleKind,
			ModuleAttributes: append([]ModuleAttribute(nil), b.moduleAttributes...), Parse: b.parse,
			Declarations: []Declaration{}, Procedures: []ProcedureIR{}, TypeReferences: []TypeReference{},
		},
		moduleScope: map[string]SymbolScope{},
	}
	if root != nil {
		v.visit(root, visitContext{procedure: -1, callIndex: -1})
	}
	if b.err != nil {
		return DocumentIR{}
	}
	v.finalize()
	return v.document
}

func (v *singleVisitor) visit(node *tree_sitter.Node, ctx visitContext) {
	if node == nil || v.builder.err != nil {
		return
	}
	v.visited++
	if v.visited&0xff == 0 {
		if err := v.builder.ctx.Err(); err != nil {
			v.builder.err = err
			return
		}
	}
	if isProcedureNode(node.Kind()) {
		ctx = v.enterProcedure(node, ctx)
	}
	if node.Kind() == "enum_declaration" {
		ctx.enumName = cleanIdentifier(nodeText(childByFieldOrKind(node, "name", "identifier"), v.builder.source))
		ctx.enumVisibility = visibilityOfNode(node, v.builder.source)
	}
	procedure := v.procedure(ctx.procedure)

	if node.Kind() == "variable_declaration" || node.Kind() == "const_declaration" {
		scope := ScopeModule
		if procedure != nil && ctx.inBody {
			scope = ScopeLocal
		}
		declarations := v.builder.declarations(node, scope)
		if procedure == nil {
			v.document.Declarations = append(v.document.Declarations, declarations...)
		} else if ctx.inBody {
			procedure.Declarations = append(procedure.Declarations, declarations...)
		}
	}
	if procedure == nil {
		switch node.Kind() {
		case "declare_statement", "declare_sub_statement", "declare_function_statement", "event_statement", "event_declaration", "type_declaration", "enum_declaration":
			if declaration, ok := v.builder.simpleDeclaration(node, ScopeModule); ok {
				switch node.Kind() {
				case "type_declaration":
					declaration.Kind = "type"
				case "enum_declaration":
					declaration.Kind = "enum"
				}
				v.document.Declarations = append(v.document.Declarations, declaration)
			}
		case "enum_member":
			if declaration, ok := v.builder.enumMemberDeclaration(node, ctx); ok {
				v.document.Declarations = append(v.document.Declarations, declaration)
			}
		}
	}
	v.addTypeReference(node, ctx)
	if procedure != nil && ctx.inBody {
		if kind, ok := statementKind(node); ok {
			if kind == StatementCall && (isIndexedAssignmentNode(node, v.builder.source) || isLetAssignmentNode(node, v.builder.source)) {
				kind = StatementAssignment
			}
			ctx.statementID = v.addStatement(procedure, node, ctx.statementID, ctx.branch, ctx.conditional, kind)
		}
		if isExpressionNode(node.Kind()) {
			ctx.parentExprID = v.addExpression(procedure, node, ctx)
		}
		ctx.callIndex = v.addCall(procedure, node, ctx)
		if node.Kind() == "raise_event_statement" {
			v.addRaiseEvent(procedure, node, ctx)
		}
		if node.Kind() == "identifier" && !ctx.metadata {
			v.addAccess(procedure, node, ctx)
		}
	}

	if node.Kind() == "block" && procedure != nil {
		v.visitBlock(node, ctx)
		return
	}
	if node.Kind() == "argument_list" && procedure != nil && ctx.callIndex >= 0 {
		v.visitArgumentList(node, ctx)
		return
	}
	if node.Kind() == "named_argument" && procedure != nil && ctx.callIndex >= 0 {
		v.visitNamedArgument(node, ctx)
		return
	}

	for i := uint(0); i < node.NamedChildCount(); i++ {
		if v.builder.err != nil {
			return
		}
		child := node.NamedChild(i)
		if child == nil {
			continue
		}
		childCtx := v.childContext(node, child, ctx)
		v.visit(child, childCtx)
	}
}

func (v *singleVisitor) enterProcedure(node *tree_sitter.Node, parent visitContext) visitContext {
	symbol := v.builder.procedureSymbol(node)
	if len(parent.conditional) > 0 {
		symbol.ConditionalBranches = append(symbol.ConditionalBranches, parent.conditional...)
	}
	procedure := ProcedureIR{
		Symbol: symbol, Declarations: []Declaration{}, Statements: []Statement{},
		Expressions: []Expression{}, Calls: []CallSite{}, Accesses: []VariableAccess{},
	}
	if symbol.Kind == ProcedureFunction || symbol.Kind == ProcedurePropertyGet || symbol.Kind == ProcedureProperty {
		nameRange := symbol.DeclarationRange
		if name := node.ChildByFieldName("name"); name != nil {
			nameRange = vbaast.NodeRange(name)
		}
		procedure.Declarations = append(procedure.Declarations, Declaration{
			ID: v.builder.takeDeclarationID(), Name: symbol.Name, Type: symbol.ReturnType,
			Scope: ScopeLocal, Kind: "return_slot", IsArray: symbol.IsArray,
			ValueShape: symbol.ValueShape, ArrayBounds: append([]ArrayBound(nil), symbol.ArrayBounds...), Range: nameRange,
		})
	}
	for _, parameter := range symbol.Parameters {
		procedure.Declarations = append(procedure.Declarations, Declaration{
			ID: v.builder.takeDeclarationID(), Name: parameter.Name, Type: parameter.Type,
			Scope: ScopeParameter, Kind: "parameter", IsArray: parameter.IsArray,
			ParamArray: parameter.ParamArray, ValueShape: valueShapeForParameter(parameter),
			ArrayBounds: append([]ArrayBound(nil), parameter.ArrayBounds...),
			IsObject:    looksObjectType(parameter.Type), Range: parameter.Range,
		})
	}
	v.document.Procedures = append(v.document.Procedures, procedure)
	v.callFacts = append(v.callFacts, []callFact{})
	return visitContext{procedure: len(v.document.Procedures) - 1, callIndex: -1, conditional: append([]ConditionalBranch(nil), parent.conditional...)}
}

func (v *singleVisitor) procedure(index int) *ProcedureIR {
	if index < 0 || index >= len(v.document.Procedures) {
		return nil
	}
	return &v.document.Procedures[index]
}

func (v *singleVisitor) addStatement(
	procedure *ProcedureIR,
	node *tree_sitter.Node,
	parentID int,
	branch BranchRole,
	conditional []ConditionalBranch,
	kind StatementKind,
) int {
	id := len(procedure.Statements) + 1
	statement := Statement{
		ID: id, ParentID: parentID, Kind: kind, SyntaxKind: node.Kind(),
		Text: nodeText(node, v.builder.source), Range: vbaast.NodeRange(node), Recovered: recovered(node),
		ConditionalBranches: append([]ConditionalBranch(nil), conditional...),
	}
	if node.Kind() == "line_number_statement" {
		if number := node.ChildByFieldName("number"); number != nil {
			statement.Text = nodeText(number, v.builder.source)
			statement.Range = vbaast.NodeRange(number)
		}
	}
	if statement.Recovered && kind == StatementUnknown {
		statement.Kind = StatementRecovered
	}
	if branch != "" {
		statement.Control = &ControlFlowMetadata{Branch: branch}
	}
	v.populateControlMetadata(&statement, node)
	switch statement.Kind {
	case StatementLabel:
		if node.Kind() == "line_number_statement" {
			label := node.ChildByFieldName("number")
			statement.Label = nodeText(label, v.builder.source)
			if label != nil {
				statement.LabelRange = vbaast.NodeRange(label)
			}
		} else {
			label := node.ChildByFieldName("name")
			statement.Label = nodeText(label, v.builder.source)
			if label != nil {
				statement.LabelRange = vbaast.NodeRange(label)
			}
		}
	case StatementGoTo, StatementOnError, StatementResume, StatementExit:
		statement.Label = statementOperand(statement.Text, statement.Kind)
		if statement.Control != nil {
			switch statement.Control.Transfer {
			case TransferGoto, TransferOnErrorGoto, TransferResumeLabel:
				statement.Label = statement.Control.Target
			}
		}
	}
	procedure.Statements = append(procedure.Statements, statement)
	return id
}

func (v *singleVisitor) populateControlMetadata(statement *Statement, node *tree_sitter.Node) {
	if statement == nil || node == nil {
		return
	}
	control := statement.Control
	ensureControl := func() *ControlFlowMetadata {
		if control == nil {
			control = &ControlFlowMetadata{}
		}
		return control
	}
	switch node.Kind() {
	case "case_clause":
		if childByKind(node, "case_expression") == nil && !recovered(node) {
			ensureControl().CaseElse = true
		}
	case "do_statement":
		if loop := doLoopTest(node, v.builder.source); loop != "" {
			ensureControl().Loop = loop
		}
	case "goto_statement":
		// GoSub/Return and computed On ... GoTo/GoSub transfers are kept as
		// conservative unknown flow. They do not have a normalized target
		// transfer in the current procedure IR and are outside VB055/VB056.
		fields := strings.Fields(nodeText(node, v.builder.source))
		if len(fields) > 0 && strings.EqualFold(fields[0], "gosub") {
			break
		}
		target := node.ChildByFieldName("target")
		if target != nil {
			value := ensureControl()
			value.Transfer = TransferGoto
			value.Target = cleanIdentifier(nodeText(target, v.builder.source))
			value.TargetRange = vbaast.NodeRange(target)
		}
	case "for_statement", "for_each_statement":
		value := ensureControl()
		if variable := node.ChildByFieldName("variable"); variable != nil {
			value.LoopVariable = cleanIdentifier(nodeText(variable, v.builder.source))
			value.LoopVariableRange = vbaast.NodeRange(variable)
		}
		if next := node.ChildByFieldName("next_variables"); next != nil && !recovered(next) {
			for i := uint(0); i < next.NamedChildCount(); i++ {
				variable := next.NamedChild(i)
				if variable == nil || recovered(variable) {
					continue
				}
				name := cleanIdentifier(nodeText(variable, v.builder.source))
				if name == "" {
					continue
				}
				value.NextVariables = append(value.NextVariables, name)
				value.NextVariableRanges = append(value.NextVariableRanges, vbaast.NodeRange(variable))
			}
		}
	case "exit_statement":
		if transfer := exitTransfer(node, v.builder.source); transfer != "" {
			ensureControl().Transfer = transfer
		}
	case "on_error_statement":
		value := ensureControl()
		if target := node.ChildByFieldName("target"); target != nil {
			value.Target = cleanIdentifier(nodeText(target, v.builder.source))
			value.TargetRange = vbaast.NodeRange(target)
			if value.Target == "0" {
				value.Transfer = TransferOnErrorDisable
				value.Target = ""
			} else {
				value.Transfer = TransferOnErrorGoto
			}
		} else {
			value.Transfer = TransferOnErrorResumeNext
		}
	case "resume_statement":
		value := ensureControl()
		if target := node.ChildByFieldName("target"); target != nil {
			value.Transfer = TransferResumeLabel
			value.Target = cleanIdentifier(nodeText(target, v.builder.source))
			value.TargetRange = vbaast.NodeRange(target)
		} else if fields := strings.Fields(nodeText(node, v.builder.source)); len(fields) == 2 &&
			strings.EqualFold(fields[0], "Resume") && strings.EqualFold(fields[1], "Next") {
			value.Transfer = TransferResumeNext
		} else {
			value.Transfer = TransferResumeRetry
		}
	case "end_statement":
		ensureControl().Transfer = TransferTerminate
	}
	if control != nil {
		control.Range = vbaast.NodeRange(node)
		statement.Control = control
	}
}

func doLoopTest(node *tree_sitter.Node, source []byte) LoopTest {
	condition := childByKind(node, "do_condition")
	if condition == nil {
		return ""
	}
	conditionText := strings.TrimSpace(nodeText(condition, source))
	until := strings.HasPrefix(strings.ToLower(conditionText), "until")
	post := condition.StartPosition().Row > node.StartPosition().Row
	if body := node.ChildByFieldName("body"); body != nil {
		post = condition.StartByte() >= body.EndByte()
	} else if int(condition.StartByte()) <= len(source) {
		prefix := string(source[node.StartByte():condition.StartByte()])
		post = hasWord(prefix, "Loop")
	}
	switch {
	case post && until:
		return LoopPostUntil
	case post:
		return LoopPostWhile
	case until:
		return LoopPreUntil
	default:
		return LoopPreWhile
	}
}

func exitTransfer(node *tree_sitter.Node, source []byte) TransferKind {
	switch statementOperand(nodeText(node, source), StatementExit) {
	case "sub":
		return TransferExitSub
	case "function":
		return TransferExitFunction
	case "property":
		return TransferExitProperty
	case "for":
		return TransferExitFor
	case "do":
		return TransferExitDo
	default:
		return ""
	}
}

func (v *singleVisitor) addExpression(procedure *ProcedureIR, node *tree_sitter.Node, ctx visitContext) int {
	id := len(procedure.Expressions) + 1
	expression := Expression{
		ID: id, ParentID: ctx.parentExprID, StatementID: ctx.statementID,
		Kind: expressionKind(node.Kind()), SyntaxKind: node.Kind(), Text: nodeText(node, v.builder.source),
		Range: vbaast.NodeRange(node), Recovered: recovered(node),
	}
	procedure.Expressions = append(procedure.Expressions, expression)
	if ctx.parentExprID > 0 {
		if parent := expressionByID(procedure, ctx.parentExprID); parent != nil {
			parent.Children = append(parent.Children, id)
		}
	} else if statement := statementByID(procedure, ctx.statementID); statement != nil {
		statement.ExpressionIDs = append(statement.ExpressionIDs, id)
	}
	if statement := statementByID(procedure, ctx.statementID); statement != nil {
		v.linkStatementExpression(statement, node, expression)
	}
	return id
}

func (v *singleVisitor) linkStatementExpression(statement *Statement, node *tree_sitter.Node, expression Expression) {
	rng := vbaast.NodeRange(node)
	if statement.Target != nil && sameRange(statement.Target.Range, rng) {
		statement.TargetID = expression.ID
	}
	if statement.Value != nil && sameRange(statement.Value.Range, rng) {
		statement.ValueID = expression.ID
	}
	if statement.Condition != nil && sameRange(statement.Condition.Range, rng) {
		statement.ConditionID = expression.ID
	}
}

func (v *singleVisitor) addCall(procedure *ProcedureIR, node *tree_sitter.Node, ctx visitContext) int {
	if ctx.suppressCall {
		return ctx.callIndex
	}
	field := ""
	switch node.Kind() {
	case "call_statement":
		field = "callee"
	case "raise_event_statement":
		field = "event"
	case "call_expression":
		field = "function"
	case "new_expression":
		target := node.ChildByFieldName("type")
		if target == nil {
			return ctx.callIndex
		}
		callee := calleeFromNode(target, v.builder.source)
		callee.Text = "New " + callee.Text
		return v.appendCall(procedure, node, callee, ctx)
	default:
		return ctx.callIndex
	}
	target := node.ChildByFieldName(field)
	if target == nil && node.Kind() == "call_statement" {
		target = firstNamedChild(node)
	}
	if target == nil {
		return ctx.callIndex
	}
	if target.Kind() == "call_expression" {
		if function := target.ChildByFieldName("function"); function != nil {
			target = function
		}
	}
	callee := calleeFromNode(target, v.builder.source)
	if callee.Text == "" {
		return ctx.callIndex
	}
	return v.appendCall(procedure, node, callee, ctx)
}

func (v *singleVisitor) addRaiseEvent(procedure *ProcedureIR, node *tree_sitter.Node, ctx visitContext) {
	if procedure == nil || node == nil {
		return
	}
	event := node.ChildByFieldName("event")
	if event == nil {
		return
	}
	name := cleanIdentifier(nodeText(event, v.builder.source))
	if name == "" {
		return
	}
	arguments, _ := v.argumentsForCall(node)
	procedure.RaiseEvents = append(procedure.RaiseEvents, RaiseEventReference{
		Name: name, Module: v.builder.moduleName,
		Caller: ProcedureRef{Name: procedure.Symbol.Name, Kind: procedure.Symbol.Kind, QualifiedName: procedure.Symbol.QualifiedName},
		Range:  vbaast.NodeRange(event), Arguments: arguments, Recovered: recovered(node),
		ConditionalBranches: append([]ConditionalBranch(nil), ctx.conditional...),
	})
}

func (v *singleVisitor) appendCall(procedure *ProcedureIR, node *tree_sitter.Node, callee Callee, ctx visitContext) int {
	expressionID := 0
	if node.Kind() == "call_expression" || node.Kind() == "new_expression" {
		expressionID = ctx.parentExprID
	}
	arguments, facts := v.argumentsForCall(node)
	procedure.Calls = append(procedure.Calls, CallSite{
		ID: len(procedure.Calls) + 1, File: v.builder.file, Module: v.builder.moduleName,
		Caller: ProcedureRef{Name: procedure.Symbol.Name, Kind: procedure.Symbol.Kind, QualifiedName: procedure.Symbol.QualifiedName},
		Callee: callee, Arguments: arguments,
		Range: vbaast.NodeRange(node), StatementID: ctx.statementID, ExpressionID: expressionID,
		IsRaiseEvent: node.Kind() == "raise_event_statement",
		Resolution:   CallResolution{Status: ResolutionNotAttempted},
	})
	procedureIndex := len(v.document.Procedures) - 1
	v.callFacts[procedureIndex] = append(v.callFacts[procedureIndex], callFact{arguments: facts})
	return len(procedure.Calls) - 1
}

func (v *singleVisitor) argumentsForCall(node *tree_sitter.Node) (Arguments, []argumentFact) {
	target := node.ChildByFieldName("function")
	if node.Kind() == "call_statement" {
		target = node.ChildByFieldName("callee")
	} else if node.Kind() == "raise_event_statement" {
		target = node.ChildByFieldName("event")
	}
	arguments := argumentsFromCallNode(node, target, v.builder.source)
	arguments.ExpressionIDs = []int{}
	list := node.ChildByFieldName("arguments")
	if list == nil && target != nil && target.Kind() == "call_expression" {
		list = target.ChildByFieldName("arguments")
	}
	if list == nil {
		list = childByKind(node, "argument_list")
	}
	if list == nil {
		for i := uint(0); i < node.NamedChildCount(); i++ {
			child := node.NamedChild(i)
			if child != nil && child.Kind() == "argument_list" {
				list = child
				break
			}
		}
	}
	if list != nil && arguments.Count == 0 {
		arguments = argumentsFromArgumentList(list, v.builder.source)
	}
	if list == nil {
		return arguments, nil
	}
	facts := make([]argumentFact, 0, list.NamedChildCount())
	for i := uint(0); i < list.NamedChildCount(); i++ {
		child := list.NamedChild(i)
		if child == nil {
			continue
		}
		fact := argumentFact{rng: vbaast.NodeRange(child)}
		if child.Kind() == "named_argument" {
			fact.name = cleanIdentifier(nodeText(child.ChildByFieldName("name"), v.builder.source))
			if value := child.ChildByFieldName("value"); value != nil {
				fact.valueText = nodeText(value, v.builder.source)
				fact.valueRange = vbaast.NodeRange(value)
			}
		}
		facts = append(facts, fact)
	}
	return arguments, facts
}

func (v *singleVisitor) addAccess(procedure *ProcedureIR, node *tree_sitter.Node, ctx visitContext) {
	name := cleanIdentifier(nodeText(node, v.builder.source))
	if name == "" {
		return
	}
	procedure.Accesses = append(procedure.Accesses, VariableAccess{
		Name: name, Mode: firstAccessMode(ctx.accessMode), Scope: ScopeUnresolved,
		Range: vbaast.NodeRange(node), StatementID: ctx.statementID, ExpressionID: ctx.parentExprID,
		Resolution: SymbolResolution{Scope: ScopeUnresolved},
	})
}

func (v *singleVisitor) addTypeReference(node *tree_sitter.Node, ctx visitContext) {
	kind, field := "", ""
	switch node.Kind() {
	case "as_type_clause":
		kind, field = "uses_type", "type"
	case "new_expression":
		kind, field = "constructs", "type"
	case "implements_statement":
		kind, field = "implements", "name"
	}
	if kind == "" {
		return
	}
	target := node.ChildByFieldName(field)
	if target == nil {
		return
	}
	ref := TypeReference{
		Kind: kind, File: v.builder.file, Module: v.builder.moduleName,
		Target: nodeText(target, v.builder.source), Range: vbaast.NodeRange(target), Parse: v.builder.parse,
	}
	if procedure := v.procedure(ctx.procedure); procedure != nil {
		caller := ProcedureRef{
			Name: procedure.Symbol.Name, Kind: procedure.Symbol.Kind, QualifiedName: procedure.Symbol.QualifiedName,
		}
		ref.Caller = &caller
	}
	v.document.TypeReferences = append(v.document.TypeReferences, ref)
	if node.Kind() == "as_type_clause" && hasWord(nodeText(node, v.builder.source), "New") {
		ref.Kind = "constructs"
		v.document.TypeReferences = append(v.document.TypeReferences, ref)
	}
}

func (v *singleVisitor) visitBlock(node *tree_sitter.Node, ctx visitContext) {
	type ifFrame struct{ outer, ifID int }
	activeParent := ctx.statementID
	stack := []ifFrame{}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		if v.builder.err != nil {
			return
		}
		child := node.NamedChild(i)
		if child == nil {
			continue
		}
		switch child.Kind() {
		case "else_fragment":
			if len(stack) == 0 {
				continue
			}
			procedure := v.procedure(ctx.procedure)
			activeParent = v.addStatement(procedure, child, stack[len(stack)-1].ifID, ctx.branch, ctx.conditional, StatementElse)
			continue
		case "end_if_fragment":
			if len(stack) > 0 {
				activeParent = stack[len(stack)-1].outer
				stack = stack[:len(stack)-1]
			}
			continue
		}
		childCtx := v.childContext(node, child, ctx)
		childCtx.statementID = activeParent
		before := len(v.procedure(ctx.procedure).Statements)
		kind, isStatement := statementKind(child)
		v.visit(child, childCtx)
		if !isStatement || len(v.procedure(ctx.procedure).Statements) == before {
			continue
		}
		id := before + 1
		switch kind {
		case StatementIf:
			if child.Kind() != "single_line_if_statement" {
				stack = append(stack, ifFrame{outer: activeParent, ifID: id})
				activeParent = id
			}
		case StatementElseIf:
			if len(stack) > 0 {
				activeParent = id
			}
		}
	}
}

func (v *singleVisitor) visitArgumentList(node *tree_sitter.Node, ctx visitContext) {
	for i := uint(0); i < node.NamedChildCount(); i++ {
		if v.builder.err != nil {
			return
		}
		child := node.NamedChild(i)
		if child == nil {
			continue
		}
		v.visit(child, v.childContext(node, child, ctx))
	}
}

func (v *singleVisitor) visitNamedArgument(node *tree_sitter.Node, ctx visitContext) {
	for i := uint(0); i < node.NamedChildCount(); i++ {
		if v.builder.err != nil {
			return
		}
		child := node.NamedChild(i)
		if child == nil {
			continue
		}
		childCtx := v.childContext(node, child, ctx)
		if sameNode(child, node.ChildByFieldName("name")) {
			childCtx.metadata = true
		}
		v.visit(child, childCtx)
	}
}

func (v *singleVisitor) childContext(parent, child *tree_sitter.Node, ctx visitContext) visitContext {
	childCtx := ctx
	if parent.Kind() == "preprocessor_if" {
		if body := parent.ChildByFieldName("body"); sameNode(child, body) {
			childCtx.conditional = appendConditionalBranch(ctx.conditional, parent, child, 0, v.builder.source)
		} else if strings.HasPrefix(child.Kind(), "preprocessor_else") {
			branch := 0
			for i := uint(0); i < parent.NamedChildCount(); i++ {
				sibling := parent.NamedChild(i)
				if sibling == nil {
					continue
				}
				if strings.HasPrefix(sibling.Kind(), "preprocessor_else") {
					branch++
				}
				if sameNode(sibling, child) {
					break
				}
			}
			childCtx.conditional = appendConditionalBranch(ctx.conditional, parent, child, branch, v.builder.source)
		}
	}
	if parent.Kind() == "single_line_if_statement" {
		switch {
		case sameNode(child, parent.ChildByFieldName("consequence")):
			childCtx.branch = BranchThen
		case sameNode(child, parent.ChildByFieldName("alternative")):
			childCtx.branch = BranchElse
		}
	}
	if isExpressionNode(parent.Kind()) {
		childCtx.parentExprID = ctx.parentExprID
	}
	if isDeclarationNode(parent.Kind()) || parent.Kind() == "as_type_clause" || parent.Kind() == "implements_statement" {
		childCtx.metadata = true
	}
	if isProcedureNode(parent.Kind()) && child.Kind() != "block" {
		childCtx.metadata = true
	}
	if isProcedureNode(parent.Kind()) && child.Kind() == "block" {
		childCtx.inBody = true
	}
	switch parent.Kind() {
	case "qualified_member_expression", "implicit_member_expression", "member_expression":
		if sameNode(child, childByFieldNameAny(parent, "member", "property")) {
			childCtx.metadata = true
		} else if isTargetAccessMode(ctx.accessMode) {
			childCtx.accessMode = AccessRead
		}
	case "new_expression":
		if sameNode(child, parent.ChildByFieldName("type")) {
			childCtx.metadata = true
		}
	case "named_argument":
		if sameNode(child, parent.ChildByFieldName("name")) {
			childCtx.metadata = true
		}
	case "raise_event_statement":
		if sameNode(child, parent.ChildByFieldName("event")) {
			childCtx.metadata = true
		}
	case "call_expression":
		if sameNode(child, parent.ChildByFieldName("function")) {
			childCtx.suppressCall = true
			if child.Kind() == "identifier" {
				if isTargetAccessMode(ctx.accessMode) {
					childCtx.metadata = false
					childCtx.accessMode = ctx.accessMode
					if statement := statementByID(v.procedure(ctx.procedure), ctx.statementID); statement != nil {
						statement.Target = expressionStub(child, ctx.statementID, v.builder.source)
					}
				} else {
					childCtx.metadata = true
				}
			}
		}
		if sameNode(child, parent.ChildByFieldName("arguments")) {
			childCtx.suppressCall = false
			if isTargetAccessMode(ctx.accessMode) {
				childCtx.accessMode = AccessRead
			}
		}
	case "call_statement":
		if sameNode(child, parent.ChildByFieldName("callee")) {
			childCtx.callIndex = ctx.callIndex
			childCtx.suppressCall = child.Kind() == "call_expression"
			if isIndexedAssignmentNode(parent, v.builder.source) && !isLetIndexedAssignmentNode(parent, v.builder.source) {
				childCtx.metadata = false
				childCtx.accessMode = AccessWrite
				if statement := statementByID(v.procedure(ctx.procedure), ctx.statementID); statement != nil {
					statement.Target = expressionStub(child, ctx.statementID, v.builder.source)
				}
			} else if child.Kind() == "identifier" {
				childCtx.metadata = true
			}
		}
		if sameNode(child, parent.ChildByFieldName("arguments")) &&
			isLetIndexedAssignmentNode(parent, v.builder.source) {
			childCtx.accessMode = AccessWrite
		}
	case "comparison_expression":
		if statement := statementByID(v.procedure(ctx.procedure), ctx.statementID); statement != nil &&
			statement.SyntaxKind == "call_statement" && statement.Kind == StatementAssignment {
			letAssignment := isDirectLetAssignmentText(statement.Text)
			if letAssignment && sameNode(child, childByFieldNameAny(parent, "left", "target")) {
				childCtx.accessMode = AccessWrite
				statement.Target = expressionStub(child, ctx.statementID, v.builder.source)
			} else {
				childCtx.accessMode = AccessRead
			}
			if sameNode(child, childByFieldNameAny(parent, "right", "value")) {
				statement.Value = expressionStub(child, ctx.statementID, v.builder.source)
			}
		}
	case "redim_statement":
		if child.Kind() == "redim_declarator" {
			if statement := statementByID(v.procedure(ctx.procedure), ctx.statementID); statement != nil {
				childCtx.accessMode = targetAccessMode(*statement)
				if target := childByKind(child, "identifier"); target != nil {
					statement.Target = expressionStub(target, ctx.statementID, v.builder.source)
				}
			}
		}
	case "redim_declarator":
		if child.Kind() == "array_bounds" {
			childCtx.accessMode = AccessRead
		}
	}
	if statement := statementByID(v.procedure(ctx.procedure), ctx.statementID); statement != nil {
		target := childByFieldNameAny(parent, "target", "left", "variable")
		value := childByFieldNameAny(parent, "value", "right")
		condition := childByFieldNameAny(parent, "condition", "test")
		indexedAssignmentComparison := parent.Kind() == "comparison_expression" &&
			statement.SyntaxKind == "call_statement" && statement.Kind == StatementAssignment
		switch {
		case sameNode(child, target) && !indexedAssignmentComparison:
			childCtx.accessMode = targetAccessMode(*statement)
			statement.Target = expressionStub(child, ctx.statementID, v.builder.source)
		case sameNode(child, value):
			childCtx.accessMode = AccessRead
			statement.Value = expressionStub(child, ctx.statementID, v.builder.source)
		case sameNode(child, condition):
			childCtx.accessMode = AccessRead
			statement.Condition = expressionStub(child, ctx.statementID, v.builder.source)
		}
	}
	return childCtx
}

func appendConditionalBranch(path []ConditionalBranch, parent, branch *tree_sitter.Node, branchNumber int, source []byte) []ConditionalBranch {
	if parent == nil || branch == nil {
		return path
	}
	condition := ""
	if branchNumber == 0 {
		condition = nodeText(parent.ChildByFieldName("condition"), source)
	} else if branch.Kind() == "preprocessor_elseif" {
		condition = nodeText(branch.ChildByFieldName("condition"), source)
	}
	return append(append([]ConditionalBranch(nil), path...), ConditionalBranch{
		Group: strconv.Itoa(int(parent.StartByte())), Condition: condition, Branch: branchNumber, Range: vbaast.NodeRange(branch),
	})
}

func (v *singleVisitor) finalize() {
	sort.SliceStable(v.document.Declarations, func(i, j int) bool {
		return v.document.Declarations[i].Range.StartByte < v.document.Declarations[j].Range.StartByte
	})
	for _, declaration := range v.document.Declarations {
		v.moduleScope[strings.ToLower(declaration.Name)] = declaration.Scope
	}
	for procedureIndex := range v.document.Procedures {
		procedure := &v.document.Procedures[procedureIndex]
		scopes := map[string]SymbolScope{}
		for name := range v.moduleScope {
			scopes[name] = ScopeModule
		}
		for _, declaration := range procedure.Declarations {
			scopes[strings.ToLower(declaration.Name)] = declaration.Scope
		}
		for statementIndex := range procedure.Statements {
			ensureIndexedAssignmentAccess(procedure, &procedure.Statements[statementIndex])
		}
		for accessIndex := range procedure.Accesses {
			access := &procedure.Accesses[accessIndex]
			if scope := scopes[strings.ToLower(access.Name)]; scope != "" {
				access.Scope = scope
				access.Resolution.Scope = scope
			}
		}
		for callIndex := range procedure.Calls {
			call := &procedure.Calls[callIndex]
			fact := &v.callFacts[procedureIndex][callIndex]
			if len(fact.arguments) > 0 {
				call.Arguments.Count = len(fact.arguments)
				call.Arguments.Named = nil
			}
			for _, argument := range fact.arguments {
				expressionID := expressionIDWithin(procedure.Expressions, argument.rng)
				call.Arguments.ExpressionIDs = append(call.Arguments.ExpressionIDs, expressionID)
				if argument.name != "" {
					valueID := expressionIDWithin(procedure.Expressions, argument.valueRange)
					call.Arguments.Named = append(call.Arguments.Named, NamedArgument{
						Name: argument.name, ValueText: argument.valueText, ExpressionID: valueID,
					})
				}
			}
			if call.ExpressionID == 0 {
				call.ExpressionID = containingExpressionID(procedure.Expressions, call.Range)
			}
		}
		for statementIndex := range procedure.Statements {
			statement := &procedure.Statements[statementIndex]
			statement.Target = canonicalExpression(procedure.Expressions, statement.TargetID, statement.Target)
			statement.Value = canonicalExpression(procedure.Expressions, statement.ValueID, statement.Value)
			statement.Condition = canonicalExpression(procedure.Expressions, statement.ConditionID, statement.Condition)
		}
		sort.SliceStable(procedure.Accesses, func(i, j int) bool {
			return procedure.Accesses[i].Range.StartByte < procedure.Accesses[j].Range.StartByte
		})
	}
	sort.SliceStable(v.document.Procedures, func(i, j int) bool {
		return v.document.Procedures[i].Symbol.DeclarationRange.StartByte < v.document.Procedures[j].Symbol.DeclarationRange.StartByte
	})
}

func statementByID(procedure *ProcedureIR, id int) *Statement {
	if procedure == nil || id <= 0 || id > len(procedure.Statements) {
		return nil
	}
	return &procedure.Statements[id-1]
}

func expressionByID(procedure *ProcedureIR, id int) *Expression {
	if procedure == nil || id <= 0 || id > len(procedure.Expressions) {
		return nil
	}
	return &procedure.Expressions[id-1]
}

func expressionStub(node *tree_sitter.Node, statementID int, source []byte) *Expression {
	if node == nil {
		return nil
	}
	return &Expression{
		StatementID: statementID, Kind: expressionKind(node.Kind()), SyntaxKind: node.Kind(),
		Text: nodeText(node, source), Range: vbaast.NodeRange(node), Recovered: recovered(node),
	}
}

func canonicalExpression(expressions []Expression, id int, fallback *Expression) *Expression {
	if id > 0 && id <= len(expressions) {
		expression := expressions[id-1]
		expression.Children = append([]int(nil), expression.Children...)
		return &expression
	}
	return fallback
}

func expressionIDWithin(expressions []Expression, rng vbaast.Range) int {
	if rng.EndByte <= rng.StartByte {
		return 0
	}
	best, width := 0, -1
	for _, expression := range expressions {
		if !containsRange(rng, expression.Range) {
			continue
		}
		candidate := expression.Range.EndByte - expression.Range.StartByte
		if candidate > width {
			best, width = expression.ID, candidate
		}
	}
	return best
}

func sameRange(a, b vbaast.Range) bool {
	return a.StartByte == b.StartByte && a.EndByte == b.EndByte
}

func firstAccessMode(mode AccessMode) AccessMode {
	if mode == "" {
		return AccessRead
	}
	return mode
}

func isTargetAccessMode(mode AccessMode) bool {
	return mode == AccessWrite || mode == AccessReadWrite
}

func isIndexedAssignmentNode(node *tree_sitter.Node, source []byte) bool {
	if node == nil || node.Kind() != "call_statement" || node.ChildByFieldName("callee") == nil {
		return false
	}
	if isLetIndexedAssignmentNode(node, source) {
		return true
	}
	callee := node.ChildByFieldName("callee")
	if callee.EndByte() >= node.EndByte() || int(node.EndByte()) > len(source) {
		return false
	}
	tail := string(source[callee.EndByte():node.EndByte()])
	return hasIndexedAssignmentTail(tail)
}

func isLetAssignmentNode(node *tree_sitter.Node, source []byte) bool {
	if node == nil || node.Kind() != "call_statement" {
		return false
	}
	return isDirectLetAssignmentText(nodeText(node, source)) || isLetIndexedAssignmentNode(node, source)
}

func isDirectLetAssignmentText(text string) bool {
	text = strings.TrimSpace(text)
	if len(text) < len("Let ") || !strings.EqualFold(text[:len("Let ")], "Let ") {
		return false
	}
	tail := strings.TrimSpace(text[len("Let "):])
	equal := strings.IndexByte(tail, '=')
	if equal <= 0 {
		return false
	}
	return !strings.ContainsAny(strings.TrimSpace(tail[:equal]), "()")
}

func isLetIndexedAssignmentNode(node *tree_sitter.Node, source []byte) bool {
	if node == nil || node.Kind() != "call_statement" {
		return false
	}
	text := nodeText(node, source)
	if len(text) < len("Let ") || !strings.EqualFold(text[:len("Let ")], "Let ") {
		return false
	}
	tail := strings.TrimSpace(text[len("Let "):])
	open := strings.IndexByte(tail, '(')
	return open > 0 && hasIndexedAssignmentTail(tail[open:])
}

func hasIndexedAssignmentTail(tail string) bool {
	if !strings.HasPrefix(tail, "(") {
		return false
	}
	depth := 0
	inString := false
	for i := 0; i < len(tail); i++ {
		switch tail[i] {
		case '"':
			if inString && i+1 < len(tail) && tail[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
		case '(':
			if !inString {
				depth++
			}
		case ')':
			if !inString && depth > 0 {
				depth--
			}
		case '=':
			if !inString && depth == 0 {
				return true
			}
		}
	}
	return false
}

func ensureIndexedAssignmentAccess(procedure *ProcedureIR, statement *Statement) {
	if procedure == nil || statement == nil || statement.Kind != StatementAssignment ||
		statement.SyntaxKind != "call_statement" {
		return
	}
	name := indexedAssignmentTargetName(statement.Text)
	if name == "" {
		return
	}
	for accessIndex := range procedure.Accesses {
		access := &procedure.Accesses[accessIndex]
		if access.StatementID == statement.ID && strings.EqualFold(access.Name, name) {
			access.Mode = AccessWrite
			return
		}
	}
	for expressionIndex := range procedure.Expressions {
		expression := &procedure.Expressions[expressionIndex]
		if expression.StatementID != statement.ID || expression.Kind != ExpressionIdentifier ||
			!strings.EqualFold(cleanIdentifier(expression.Text), name) {
			continue
		}
		statement.TargetID = expression.ID
		statement.Target = canonicalExpression(procedure.Expressions, expression.ID, nil)
		procedure.Accesses = append(procedure.Accesses, VariableAccess{
			Name: name, Mode: AccessWrite, Scope: ScopeUnresolved, Range: expression.Range,
			StatementID: statement.ID, ExpressionID: expression.ID,
			Resolution: SymbolResolution{Scope: ScopeUnresolved},
		})
		return
	}
}

func indexedAssignmentTargetName(text string) string {
	text = strings.TrimSpace(text)
	if len(text) >= len("Let ") && strings.EqualFold(text[:len("Let ")], "Let ") {
		text = strings.TrimSpace(text[len("Let "):])
	}
	open := strings.IndexByte(text, '(')
	if open <= 0 {
		return ""
	}
	return cleanIdentifier(text[:open])
}

func targetAccessMode(statement Statement) AccessMode {
	switch statement.Kind {
	case StatementAssignment, StatementSet, StatementForEach, StatementReDim:
		if statement.Kind == StatementReDim && hasWord(statement.Text, "Preserve") {
			return AccessReadWrite
		}
		return AccessWrite
	case StatementFor:
		return AccessReadWrite
	default:
		return AccessRead
	}
}

func isDeclarationNode(kind string) bool {
	switch kind {
	case "variable_declaration", "const_declaration", "parameter", "declare_statement",
		"declare_sub_statement", "declare_function_statement", "event_statement", "event_declaration":
		return true
	default:
		return false
	}
}
