package analyze

import (
	"regexp"
	"sort"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// unusedUDTMemberFindings reports members of Private Type declarations that no
// resolvable member expression in the module ever touches. Private types are
// file-local, so member usage is fully observable within the module. A member
// is reported only when every access path is accounted for: unresolved
// receivers mark matching member names as used, and a user-defined type that
// escapes through ByRef calls, Variants, serialization statements, or dynamic
// member strings is exempted entirely.
func (a Analyzer) unusedUDTMemberFindings(file parsedFile, signatures map[string]procedureSignature) []Finding {
	if file.IR.Parse.HasError || file.IR.Parse.HasMissing {
		return nil
	}
	types := parsePrivateUDTs(file)
	if len(types) == 0 {
		return nil
	}

	usage := &udtMemberUsage{types: types, signatures: signatures, file: file}
	usage.collectModuleEnvironment(&file.IR)
	for _, procedure := range file.IR.Procedures {
		usage.observeProcedure(procedure)
	}
	usage.observeModuleLevelText(file)

	var findings []Finding
	names := make([]string, 0, len(types))
	for name := range types {
		names = append(names, name)
	}
	sort.Strings(names)
	for _, name := range names {
		typ := types[name]
		if typ.escaped || typ.broken {
			continue
		}
		for _, memberName := range typ.order {
			member := typ.members[memberName]
			if member.used {
				continue
			}
			finding := a.simpleFinding(file, sourceProcedure{}, member.rng.StartLine, "VBA262", "information",
				"Member "+member.display+" of Private Type "+typ.display+" is never accessed.",
				"No resolvable member expression in this module reads or writes the member.",
				"Remove the member, or access it where the structure requires it.")
			finding.Column = member.rng.StartColumn
			finding.EndLine = member.rng.EndLine
			finding.EndColumn = member.rng.EndColumn
			findings = append(findings, finding)
		}
	}
	return findings
}

type udtMemberInfo struct {
	name     string
	display  string
	typeName string // canonical member type name
	rng      vbaast.Range
	used     bool
}

type udtTypeInfo struct {
	name    string
	display string
	members map[string]*udtMemberInfo
	order   []string
	escaped bool // at least one value may flow through an unmodeled boundary
	broken  bool // member declarations could not be modeled completely
}

type udtMemberUsage struct {
	types      map[string]*udtTypeInfo
	signatures map[string]procedureSignature
	file       parsedFile
	moduleVars map[string]string // canonical variable name → canonical type name
}

var (
	udtTypeEndRE     = regexp.MustCompile(`(?i)^\s*end\s+type\b`)
	udtMemberDeclRE  = regexp.MustCompile(`(?i)^\s*([a-z_][a-z0-9_]*)\s*(\([^)]*\))?\s+as\s+([a-z_][a-z0-9_.]*)`)
	udtConditionalRE = regexp.MustCompile(`^\s*#`)
)

// parsePrivateUDTs extracts Private Type member declarations from the module
// text inside the ranges the IR records for type declarations. Member lines
// the simple shape cannot parse mark the whole type as broken instead of
// guessing.
func parsePrivateUDTs(file parsedFile) map[string]*udtTypeInfo {
	types := make(map[string]*udtTypeInfo)
	for _, declaration := range file.IR.Declarations {
		if !strings.EqualFold(declaration.Kind, "type") ||
			!strings.EqualFold(strings.TrimSpace(declaration.Visibility), "private") ||
			declaration.Recovered || len(declaration.ConditionalBranches) > 0 {
			continue
		}
		name := assignmentCanonicalName(declaration.Name)
		if name == "" {
			continue
		}
		typ := &udtTypeInfo{name: name, display: cleanIdentifier(declaration.Name), members: map[string]*udtMemberInfo{}}
		start := declaration.Range.StartLine - 1
		end := declaration.Range.EndLine
		if start < 0 {
			start = 0
		}
		if end > len(file.Lines) {
			end = len(file.Lines)
		}
		for i := start + 1; i < end; i++ {
			line := file.Lines[i]
			if udtConditionalRE.MatchString(line) {
				typ.broken = true
				continue
			}
			trimmed := strings.TrimSpace(line)
			if trimmed == "" || strings.HasPrefix(trimmed, "'") || strings.HasPrefix(strings.ToLower(trimmed), "rem ") || udtTypeEndRE.MatchString(trimmed) {
				continue
			}
			match := udtMemberDeclRE.FindStringSubmatchIndex(line)
			if match == nil {
				typ.broken = true
				continue
			}
			memberName := assignmentCanonicalName(line[match[2]:match[3]])
			typeName := assignmentCanonicalName(line[match[6]:match[7]])
			if memberName == "" {
				typ.broken = true
				continue
			}
			if _, exists := typ.members[memberName]; exists {
				typ.broken = true
				continue
			}
			typ.members[memberName] = &udtMemberInfo{
				name: memberName, display: line[match[2]:match[3]], typeName: typeName,
				rng: vbaast.Range{
					StartLine: i + 1, StartColumn: match[2] + 1, StartByte: -1,
					EndLine: i + 1, EndColumn: match[3] + 1, EndByte: -1,
				},
			}
			typ.order = append(typ.order, memberName)
		}
		if len(typ.members) == 0 {
			typ.broken = true
		}
		types[name] = typ
	}
	return types
}

func (u *udtMemberUsage) collectModuleEnvironment(document *procedureir.DocumentIR) {
	u.moduleVars = make(map[string]string)
	for _, declaration := range document.Declarations {
		if declaration.Scope != procedureir.ScopeModule && declaration.Scope != "" {
			continue
		}
		if typeName := u.udtTypeName(declaration.Type); typeName != "" {
			u.moduleVars[assignmentCanonicalName(declaration.Name)] = typeName
		}
	}
}

// observeProcedure marks member usage for one procedure. The environment is
// the module-level set overlaid with parameters, local declarations, and the
// function return slot.
func (u *udtMemberUsage) observeProcedure(procedure procedureir.ProcedureIR) {
	if procedure.Symbol.Recovered || len(procedure.Symbol.ConditionalBranches) > 0 {
		u.escapeAll()
		return
	}
	env := make(map[string]string, len(u.moduleVars)+len(procedure.Declarations)+len(procedure.Symbol.Parameters)+1)
	for name, typeName := range u.moduleVars {
		env[name] = typeName
	}
	for _, declaration := range procedure.Declarations {
		if declaration.Scope != procedureir.ScopeLocal {
			continue
		}
		if typeName := u.udtTypeName(declaration.Type); typeName != "" {
			env[assignmentCanonicalName(declaration.Name)] = typeName
		}
	}
	for _, parameter := range procedure.Symbol.Parameters {
		if typeName := u.udtTypeName(parameter.Type); typeName != "" {
			env[assignmentCanonicalName(parameter.Name)] = typeName
		}
	}
	if typeName := u.udtTypeName(procedure.Symbol.ReturnType); typeName != "" {
		env[assignmentCanonicalName(procedure.Symbol.Name)] = typeName
	}
	// Even with no UDT-typed variable in scope, unresolved member expressions
	// (late-bound Variant/Object receivers) must still mark matching member
	// names used — that is the fail-open direction for dynamic access.

	statements := make(map[int]procedureir.Statement, len(procedure.Statements))
	for _, statement := range procedure.Statements {
		statements[statement.ID] = statement
	}
	withReceiver := make(map[int]int) // statement ID → With receiver expression ID
	withStatements := make([]procedureir.Statement, 0)
	for _, statement := range procedure.Statements {
		if statement.Kind == procedureir.StatementWith {
			withStatements = append(withStatements, statement)
		}
	}
	// Process With statements in source order so a nested With overwrites the
	// receiver mapping for its own subtree.
	sort.Slice(withStatements, func(i, j int) bool {
		return withStatements[i].Range.StartByte < withStatements[j].Range.StartByte
	})
	for _, statement := range withStatements {
		receiver := udtWithReceiverExpression(procedure, statement)
		if receiver == 0 {
			continue
		}
		for _, descendant := range udtDescendantStatements(procedure, statement) {
			withReceiver[descendant] = receiver
		}
	}

	// Escape: a UDT-typed variable passed where a callee may write it, or
	// touched by a statement whose write shape is unmodeled.
	writeExpr := make(map[int]bool)
	accessesByStatement := make(map[int][]procedureir.VariableAccess)
	expressionByID := make(map[int]procedureir.Expression, len(procedure.Expressions))
	for _, expression := range procedure.Expressions {
		expressionByID[expression.ID] = expression
	}
	for _, access := range procedure.Accesses {
		accessesByStatement[access.StatementID] = append(accessesByStatement[access.StatementID], access)
	}
	for _, call := range procedure.Calls {
		for id := range callWritableArgumentExpressions(call, u.signatures) {
			writeExpr[id] = true
		}
	}
	for _, access := range procedure.Accesses {
		name := assignmentCanonicalName(access.Name)
		typeName, ok := env[name]
		if !ok {
			continue
		}
		statement, exists := statements[access.StatementID]
		if !exists || assignmentAmbiguousStatement(statement) || writeExpr[access.ExpressionID] {
			u.escape(typeName)
			continue
		}
		if statement.Kind == procedureir.StatementAssignment || statement.Kind == procedureir.StatementSet {
			if udtAssignmentEscapes(statement, access, env, typeName) {
				u.escape(typeName)
			}
		}
	}

	// Member usage: every member_access expression resolves its receiver
	// through the type environment, With blocks, and call return types.
	for _, expression := range procedure.Expressions {
		if expression.Kind != procedureir.ExpressionMember {
			continue
		}
		u.observeMemberExpression(procedure, expression, expressionByID, statements, withReceiver, env)
	}
}

func (u *udtMemberUsage) udtTypeName(typeName string) string {
	canonical := assignmentCanonicalName(typeName)
	if _, ok := u.types[canonical]; ok {
		return canonical
	}
	return ""
}

func (u *udtMemberUsage) escape(typeName string) {
	if typ, ok := u.types[typeName]; ok {
		typ.escaped = true
	}
}

func (u *udtMemberUsage) escapeAll() {
	for _, typ := range u.types {
		typ.escaped = true
	}
}

// observeMemberExpression resolves a qualified or implicit member expression.
// A receiver resolving to a private UDT marks the member; an unresolved
// receiver marks the member name on every private UDT that declares it, which
// is the fail-open direction for dynamic or Variant receivers.
func (u *udtMemberUsage) observeMemberExpression(procedure procedureir.ProcedureIR, expression procedureir.Expression, expressionByID map[int]procedureir.Expression, statements map[int]procedureir.Statement, withReceiver map[int]int, env map[string]string) {
	memberName, receiverID := udtMemberParts(expression, expressionByID)
	if memberName == "" {
		return
	}
	receiverType, resolved := "", false
	if receiverID != 0 {
		receiverType, resolved = u.resolveExpressionType(procedure, receiverID, expressionByID, statements, withReceiver, env, 0)
	} else if receiver, ok := withReceiver[expression.StatementID]; ok {
		receiverType, resolved = u.resolveExpressionType(procedure, receiver, expressionByID, statements, withReceiver, env, 0)
	}
	if resolved {
		u.markMember(receiverType, memberName)
		return
	}
	u.markMemberName(memberName)
}

func (u *udtMemberUsage) markMember(typeName, memberName string) {
	typ, ok := u.types[typeName]
	if !ok {
		return
	}
	if member, ok := typ.members[memberName]; ok {
		member.used = true
	}
}

func (u *udtMemberUsage) markMemberName(memberName string) {
	for _, typ := range u.types {
		if member, ok := typ.members[memberName]; ok {
			member.used = true
		}
	}
}

// resolveExpressionType returns the private-UDT type name an expression
// denotes, and whether the receiver chain resolved far enough to attribute
// member access. ok=false means the member cannot be attributed to a known
// UDT and callers must fail open on the member name.
func (u *udtMemberUsage) resolveExpressionType(procedure procedureir.ProcedureIR, expressionID int, expressionByID map[int]procedureir.Expression, statements map[int]procedureir.Statement, withReceiver map[int]int, env map[string]string, depth int) (string, bool) {
	if depth > 16 {
		return "", false
	}
	expression, ok := expressionByID[expressionID]
	if !ok {
		return "", false
	}
	switch expression.Kind {
	case procedureir.ExpressionIdentifier:
		text := strings.TrimSpace(expression.Text)
		if bang := strings.IndexByte(text, '!'); bang > 0 {
			receiver := assignmentCanonicalName(text[:bang])
			member := assignmentCanonicalName(text[bang+1:])
			if member != "" {
				if typeName, ok := env[receiver]; ok {
					u.markMember(typeName, member)
				} else {
					u.markMemberName(member)
				}
			}
			return "", false
		}
		if typeName, ok := env[assignmentCanonicalName(text)]; ok {
			return typeName, true
		}
		return "", false
	case procedureir.ExpressionMember:
		memberName, receiverID := udtMemberParts(expression, expressionByID)
		receiverType, resolved := "", false
		if receiverID != 0 {
			receiverType, resolved = u.resolveExpressionType(procedure, receiverID, expressionByID, statements, withReceiver, env, depth+1)
		} else if receiver, ok := withReceiver[expression.StatementID]; ok {
			receiverType, resolved = u.resolveExpressionType(procedure, receiver, expressionByID, statements, withReceiver, env, depth+1)
		}
		if !resolved {
			if memberName != "" {
				u.markMemberName(memberName)
			}
			return "", false
		}
		typ, isUDT := u.types[receiverType]
		if !isUDT {
			return "", false
		}
		member, found := typ.members[memberName]
		if !found {
			return "", false
		}
		member.used = true
		return member.typeName, true
	case procedureir.ExpressionCall:
		name := assignmentCanonicalName(udtCallFunctionName(expression.Text))
		if name == "" {
			return "", false
		}
		if strings.EqualFold(name, assignmentCanonicalName(procedure.Symbol.Name)) {
			if typeName := u.udtTypeName(procedure.Symbol.ReturnType); typeName != "" {
				return typeName, true
			}
			return "", false
		}
		for _, key := range []string{name, strings.ToLower(strings.TrimSpace(u.file.IR.ModuleName) + "." + name)} {
			if signature, ok := u.signatures[key]; ok {
				if typeName := u.udtTypeName(signature.ReturnType); typeName != "" {
					return typeName, true
				}
				return "", false
			}
		}
		return "", false
	case procedureir.ExpressionParentheses:
		for _, child := range expression.Children {
			if typeName, ok := u.resolveExpressionType(procedure, child, expressionByID, statements, withReceiver, env, depth+1); ok {
				return typeName, true
			}
		}
		return "", false
	default:
		return "", false
	}
}

// udtMemberParts splits a member-access expression into the member identifier
// and the receiver expression ID. The receiver is the first child expression
// (identifier, member chain, or call); the member is the last identifier
// child. Implicit members (`.X` inside With) have no receiver child and
// return zero.
func udtMemberParts(expression procedureir.Expression, expressionByID map[int]procedureir.Expression) (string, int) {
	if len(expression.Children) == 0 {
		return "", 0
	}
	last, ok := expressionByID[expression.Children[len(expression.Children)-1]]
	if !ok || last.Kind != procedureir.ExpressionIdentifier {
		return "", 0
	}
	member := assignmentCanonicalName(last.Text)
	if member == "" {
		return "", 0
	}
	receiver := 0
	if len(expression.Children) >= 2 {
		receiver = expression.Children[0]
	}
	return member, receiver
}

func udtWithReceiverExpression(procedure procedureir.ProcedureIR, statement procedureir.Statement) int {
	if statement.Value != nil {
		return statement.Value.ID
	}
	for _, expression := range procedure.Expressions {
		if expression.StatementID == statement.ID && expression.ParentID == 0 {
			return expression.ID
		}
	}
	return 0
}

func udtDescendantStatements(procedure procedureir.ProcedureIR, withStatement procedureir.Statement) []int {
	children := make(map[int][]int)
	for _, statement := range procedure.Statements {
		if statement.ParentID != 0 {
			children[statement.ParentID] = append(children[statement.ParentID], statement.ID)
		}
	}
	var out []int
	stack := []int{withStatement.ID}
	for len(stack) > 0 {
		id := stack[len(stack)-1]
		stack = stack[:len(stack)-1]
		for _, child := range children[id] {
			out = append(out, child)
			stack = append(stack, child)
		}
	}
	return out
}

// udtAssignmentEscapes reports whether a UDT-typed variable read in an
// assignment flows into a target that is not provably the same UDT type (a
// Variant, object, or unknown target may alias or serialize it).
func udtAssignmentEscapes(statement procedureir.Statement, access procedureir.VariableAccess, env map[string]string, accessType string) bool {
	if access.Mode != procedureir.AccessRead && access.Mode != procedureir.AccessReadWrite {
		return false
	}
	if statement.Target == nil || statement.Target.Kind != procedureir.ExpressionIdentifier {
		return false
	}
	targetName := assignmentCanonicalName(statement.Target.Text)
	if targetName == "" {
		return false
	}
	targetType, ok := env[targetName]
	if !ok {
		return true
	}
	return targetType != accessType
}

// udtCallFunctionName extracts the callee base name from a call expression
// text such as `F()`, `Mod.F(x)`, or `obj.Helper()`. The last member segment
// is the function name; the receiver chain is resolved separately when the
// callee is qualified.
func udtCallFunctionName(text string) string {
	head := strings.TrimSpace(text)
	if open := strings.IndexByte(head, '('); open >= 0 {
		head = head[:open]
	}
	if index := strings.LastIndexAny(head, ".!"); index >= 0 {
		head = head[index+1:]
	}
	return strings.TrimSpace(head)
}

func (u *udtMemberUsage) observeModuleLevelText(file parsedFile) {
	// String literals may carry member names for CallByName or serialization
	// helpers; a matching member name on any private UDT counts as used.
	for _, line := range file.Lines {
		for i := 0; i < len(line); i++ {
			if line[i] != '"' {
				continue
			}
			j := i + 1
			for j < len(line) {
				if line[j] == '"' {
					if j+1 < len(line) && line[j+1] == '"' {
						j += 2
						continue
					}
					break
				}
				j++
			}
			if j >= len(line) {
				break
			}
			u.markStringLiteralMembers(line[i+1 : j])
			i = j
		}
	}
}

func (u *udtMemberUsage) markStringLiteralMembers(literal string) {
	for _, typ := range u.types {
		for memberName, member := range typ.members {
			if member.used {
				continue
			}
			if stringLiteralContainsName(literal, memberName) {
				member.used = true
			}
		}
	}
}
