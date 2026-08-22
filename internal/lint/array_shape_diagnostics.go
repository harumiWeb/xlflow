package lint

import (
	"strconv"
	"strings"
	"unicode"

	"github.com/harumiWeb/xlflow/internal/typedb"
	"github.com/harumiWeb/xlflow/internal/vba/constexpr"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// arrayShapeIssues performs the two small, compile-equivalent checks that are
// independent of runtime allocation state: writing a Const and a fixed array
// declarator whose constant lower bound exceeds its upper bound.
func (l Linter) arrayShapeIssues(path, source string, ir *procedureir.DocumentIR) []Issue {
	if ir == nil || ir.Parse.HasError || ir.Parse.HasMissing {
		return nil
	}
	consts := map[string]bool{}
	qualifiedConsts := map[string]bool{}
	constValues := map[string]int{}
	constExprs := map[string]string{}
	for name, value := range l.ConstantValues {
		if value.Kind != constexpr.ValueInteger && value.Kind != constexpr.ValueLong && value.Kind != constexpr.ValueLongLong {
			continue
		}
		integer, ok := constexpr.IntegerAsInt(value)
		if !ok {
			continue
		}
		key := normalizeQualifiedIdentifier(name)
		if key == "" {
			continue
		}
		if !strings.Contains(key, ".") && l.VisibleConstants != nil && !l.VisibleConstants[key] {
			continue
		}
		constValues[key] = integer
		consts[key] = true
		if strings.Contains(key, ".") {
			qualifiedConsts[key] = true
		}
	}
	addQualifiedConst := func(qualifier, name string) {
		qualifier = normalizeQualifiedIdentifier(qualifier)
		name = normalizeQualifiedIdentifier(name)
		if qualifier == "" || name == "" {
			return
		}
		qualifiedConsts[qualifier+"."+name] = true
	}
	lines := strings.Split(strings.ReplaceAll(source, "\r\n", "\n"), "\n")
	for name := range l.VisibleConstants {
		key := normalizeQualifiedIdentifier(name)
		if strings.Contains(key, ".") {
			qualifiedConsts[key] = true
		} else if key != "" {
			consts[key] = true
		}
	}
	inEnum := false
	var enumCurrent *int
	for lineNo, line := range lines {
		if arrayShapeLineInProcedure(ir, lineNo+1) {
			// Procedure-local Const values are resolved per procedure below;
			// keeping them out of the module table prevents cross-procedure
			// shadowing and stale values from leaking into bound checks.
			continue
		}
		trim := strings.TrimSpace(strings.SplitN(line, "'", 2)[0])
		lower := strings.ToLower(trim)
		if strings.HasPrefix(lower, "enum ") || strings.HasPrefix(lower, "public enum ") || strings.HasPrefix(lower, "private enum ") || strings.HasPrefix(lower, "friend enum ") {
			inEnum = true
			enumCurrent = nil
			continue
		}
		if inEnum && strings.HasPrefix(lower, "end enum") {
			inEnum = false
			enumCurrent = nil
			continue
		}
		if inEnum && trim != "" {
			name, expression := enumMemberConstant(trim, enumCurrent)
			if name != "" {
				key := strings.ToLower(cleanIdentifier(name))
				consts[key] = true
				constExprs[key] = expression
				if value, ok := parseArrayConstant(expression, consts, constValues); ok {
					constValues[key] = value
					next := value + 1
					enumCurrent = &next
				} else {
					enumCurrent = nil
				}
			}
			continue
		}
		if name, expression, ok := arrayShapeConstLine(trim); ok {
			key := strings.ToLower(cleanIdentifier(name))
			consts[key] = true
			constExprs[key] = expression
		}
	}
	for pass := 0; pass < len(constExprs); pass++ {
		changed := false
		for name, expr := range constExprs {
			if value, ok := parseArrayConstant(expr, consts, constValues); ok {
				if old, exists := constValues[name]; !exists || old != value {
					constValues[name] = value
					changed = true
				}
			}
		}
		if !changed {
			break
		}
	}
	for _, d := range ir.Declarations {
		if strings.EqualFold(d.Kind, "const") || strings.EqualFold(d.Kind, "enum_member") {
			name := strings.ToLower(cleanIdentifier(d.Name))
			consts[name] = true
			addQualifiedConst(ir.ModuleName, d.Name)
			addQualifiedConst(d.Parent, d.Name)
			if strings.EqualFold(ir.ModuleKind, "class") || strings.EqualFold(ir.ModuleKind, "form") || strings.EqualFold(ir.ModuleKind, "document") {
				addQualifiedConst("Me", d.Name)
			}
		}
	}
	for _, p := range ir.Procedures {
		for _, d := range p.Declarations {
			if strings.EqualFold(d.Kind, "const") || strings.EqualFold(d.Kind, "enum_member") {
				consts[strings.ToLower(cleanIdentifier(d.Name))] = true
			}
		}
	}
	// TypeLib enum members/constants are part of the resolved symbol space even
	// though they do not appear in the source IR. A missing or incomplete
	// runtime database is deliberately fail-open.
	// Batch and LSP adapters supply a project-visible table that already includes
	// the immutable runtime symbols. Only the standalone adapter needs to load
	// TypeDB here; avoid reloading it once per document in project scans.
	if l.VisibleConstants == nil {
		if typeDBResult, err := typedb.LoadForRuntime(""); err == nil && typeDBResult.DB != nil {
			typeLibCounts := map[string]int{}
			constants := typeDBResult.DB.AllConstantsList()
			for _, constant := range constants {
				key := normalizeQualifiedIdentifier(constant.Name)
				if key != "" {
					typeLibCounts[key]++
				}
				addQualifiedConst(constant.EnumGroup, constant.Name)
				addQualifiedConst(constant.Library, constant.Name)
			}
			for _, constant := range constants {
				key := normalizeQualifiedIdentifier(constant.Name)
				if typeLibCounts[key] != 1 {
					delete(consts, key)
					continue
				}
				if _, sourceExists := consts[key]; sourceExists {
					// A source Const and a TypeLib constant with the same unqualified
					// name are ambiguous without an explicit qualifier.
					delete(consts, key)
					continue
				}
				consts[key] = true
			}
		}
	}
	var issues []Issue
	for _, p := range ir.Procedures {
		for _, stmt := range p.Statements {
			if stmt.Recovered || (stmt.Kind != procedureir.StatementAssignment && stmt.Kind != procedureir.StatementSet) || stmt.Target == nil {
				continue
			}
			if isConstAssignmentTarget(stmt.Target.Text, stmt.Target.Kind, p, ir.Declarations, consts, qualifiedConsts) {
				issue := l.issueAt(path, stmt.Target.Range, "VB060", "error", "A Const value cannot be assigned.")
				issue.EndLine, issue.EndColumn = stmt.Target.Range.EndLine, stmt.Target.Range.EndColumn
				issue.Kind, issue.Symbol = "constant_assignment", stmt.Target.Text
				issues = append(issues, issue)
			}
		}
	}
	procedureHeaderLines := make(map[int]bool, len(ir.Procedures))
	for _, procedure := range ir.Procedures {
		line := procedure.Symbol.DeclarationRange.StartLine
		if line > 0 {
			procedureHeaderLines[line] = true
		}
	}
	optionBase := 0
	for _, line := range lines {
		fields := strings.Fields(strings.ToLower(strings.TrimSpace(strings.SplitN(line, "'", 2)[0])))
		if len(fields) >= 3 && fields[0] == "option" && fields[1] == "base" {
			if value, err := strconv.Atoi(fields[2]); err == nil && (value == 0 || value == 1) {
				optionBase = value
			}
		}
	}
	for lineNo, line := range lines {
		if procedureHeaderLines[lineNo+1] {
			continue
		}
		trim := strings.TrimSpace(strings.SplitN(line, "'", 2)[0])
		lower := strings.ToLower(trim)
		if trim == "" || strings.HasPrefix(lower, "const ") || strings.Contains(lower, " const ") || strings.HasPrefix(lower, "redim ") || strings.Contains(lower, " redim ") {
			continue
		}
		if !declarationPrefix(lower) {
			continue
		}
		body := declarationBody(trim)
		scopeValues := arrayShapeConstantsForLine(lineNo+1, lines, ir, constValues)
		for _, part := range splitArrayDeclarators(body) {
			name, bounds, ok := arrayDeclaratorBounds(part)
			if !ok || strings.TrimSpace(bounds) == "" {
				continue
			}
			for _, dim := range splitArrayDeclarators(bounds) {
				pieces := strings.SplitN(strings.ToLower(dim), " to ", 2)
				lo, lok := optionBase, true
				upperExpr := strings.TrimSpace(dim)
				if len(pieces) == 2 {
					lo, lok = evalArrayConstant(strings.TrimSpace(pieces[0]), nil, scopeValues)
					upperExpr = strings.TrimSpace(pieces[1])
				}
				hi, hik := evalArrayConstant(upperExpr, nil, scopeValues)
				if lok && hik && lo > hi {
					issue := l.issue(path, lineNo+1, "VB061", "error", "Array lower bound cannot exceed its upper bound.")
					issue.Kind, issue.Symbol = "array_bounds", name
					issues = append(issues, issue)
					break
				}
			}
		}
	}
	return issues
}

func arrayShapeLineInProcedure(ir *procedureir.DocumentIR, line int) bool {
	if ir == nil || line < 1 {
		return false
	}
	for _, procedure := range ir.Procedures {
		range_ := procedure.Symbol.DeclarationRange
		if line >= range_.StartLine && line <= range_.EndLine {
			return true
		}
	}
	return false
}

// arrayShapeConstantsForLine returns the integer constants visible at one
// declaration line. Module constants seed the scope; procedure-local Consts
// override them, while parameters/variables shadow the same names and remove
// them from the evaluator. Unresolved local constants remain unknown rather
// than inheriting an outer value.
func arrayShapeConstantsForLine(line int, lines []string, ir *procedureir.DocumentIR, moduleValues map[string]int) map[string]int {
	values := make(map[string]int, len(moduleValues))
	for name, value := range moduleValues {
		values[name] = value
	}
	if ir == nil {
		return values
	}
	var procedure *procedureir.ProcedureIR
	for i := range ir.Procedures {
		candidate := &ir.Procedures[i]
		if line >= candidate.Symbol.DeclarationRange.StartLine && line <= candidate.Symbol.DeclarationRange.EndLine {
			procedure = candidate
			break
		}
	}
	if procedure == nil {
		return values
	}
	localNames := make(map[string]bool)
	localConst := make(map[string]bool)
	for _, declaration := range procedure.Declarations {
		key := strings.ToLower(cleanIdentifier(declaration.Name))
		if key == "" {
			continue
		}
		if strings.EqualFold(declaration.Kind, "const") || strings.EqualFold(declaration.Kind, "enum_member") {
			localConst[key] = true
			localNames[key] = true
			continue
		}
		if declaration.Scope == procedureir.ScopeLocal || declaration.Scope == procedureir.ScopeParameter {
			localNames[key] = true
		}
	}
	for name := range localNames {
		delete(values, name)
	}
	expressions := make(map[string]string)
	start := max(1, procedure.Symbol.DeclarationRange.StartLine)
	end := min(len(lines), procedure.Symbol.DeclarationRange.EndLine)
	for lineNo := start; lineNo <= end; lineNo++ {
		name, expression, ok := arrayShapeConstLine(lines[lineNo-1])
		key := strings.ToLower(cleanIdentifier(name))
		if ok && localConst[key] {
			expressions[key] = expression
		}
	}
	for pass := 0; pass < len(expressions); pass++ {
		changed := false
		for name, expression := range expressions {
			if value, ok := parseArrayConstant(expression, nil, values); ok {
				if old, exists := values[name]; !exists || old != value {
					values[name] = value
					changed = true
				}
			}
		}
		if !changed {
			break
		}
	}
	return values
}

func arrayShapeConstLine(line string) (string, string, bool) {
	trim := strings.TrimSpace(strings.SplitN(line, "'", 2)[0])
	lower := strings.ToLower(trim)
	for _, prefix := range []string{"const ", "public const ", "private const ", "friend const ", "static const "} {
		if !strings.HasPrefix(lower, prefix) {
			continue
		}
		rest := strings.TrimSpace(trim[len(prefix):])
		eq := strings.Index(rest, "=")
		if eq <= 0 {
			return "", "", false
		}
		fields := strings.Fields(strings.TrimSpace(rest[:eq]))
		if len(fields) == 0 {
			return "", "", false
		}
		return fields[0], strings.TrimSpace(rest[eq+1:]), true
	}
	return "", "", false
}

func enumMemberConstant(line string, previous *int) (string, string) {
	line = strings.TrimSpace(line)
	if line == "" {
		return "", ""
	}
	parts := strings.SplitN(line, "=", 2)
	name := strings.TrimSpace(parts[0])
	if name == "" {
		return "", ""
	}
	if fields := strings.Fields(name); len(fields) > 0 {
		name = fields[0]
	}
	if len(parts) == 2 {
		return name, strings.TrimSpace(parts[1])
	}
	if previous == nil {
		return name, "0"
	}
	return name, strconv.Itoa(*previous)
}

func isConstAssignmentTarget(name string, targetKind procedureir.ExpressionKind, procedure procedureir.ProcedureIR, moduleDeclarations []procedureir.Declaration, consts, qualifiedConsts map[string]bool) bool {
	rawName := strings.TrimSpace(name)
	if targetKind == procedureir.ExpressionMember && strings.HasPrefix(rawName, ".") {
		// Preserve the omitted receiver before normalization. This also covers
		// nested implicit chains such as `.Modes.ModeBad`, whose normalized form
		// would otherwise look like an explicit qualified constant reference.
		return false
	}
	name = normalizeQualifiedIdentifier(name)
	if targetKind != procedureir.ExpressionIdentifier && targetKind != procedureir.ExpressionMember {
		return false
	}
	// An implicit member expression (for example, `.Hidden` inside a With
	// block) is normalized to its final member name. That name alone is not
	// enough to prove a Const assignment: it may be a writable property on the
	// With receiver, or any other late-bound member. Only a complete member
	// chain can be compared with a qualified constant name.
	if targetKind == procedureir.ExpressionMember && !strings.Contains(name, ".") {
		return false
	}
	if strings.Contains(name, ".") {
		return qualifiedConsts[name]
	}
	for _, declaration := range procedure.Declarations {
		if !strings.EqualFold(cleanIdentifier(declaration.Name), name) {
			continue
		}
		return procedureir.IsConstKind(declaration.Kind)
	}
	for _, declaration := range moduleDeclarations {
		if !strings.EqualFold(cleanIdentifier(declaration.Name), name) {
			continue
		}
		return procedureir.IsConstKind(declaration.Kind)
	}
	return consts[name]
}

func normalizeQualifiedIdentifier(text string) string {
	parts := strings.Split(strings.TrimSpace(text), ".")
	for i := range parts {
		parts[i] = strings.ToLower(cleanIdentifier(strings.TrimSpace(parts[i])))
	}
	for len(parts) > 0 && parts[0] == "" {
		parts = parts[1:]
	}
	for len(parts) > 0 && parts[len(parts)-1] == "" {
		parts = parts[:len(parts)-1]
	}
	return strings.Join(parts, ".")
}

// CompileEquivalentArrayShapeIssues exposes the small, parser/IR-backed
// compile-equivalent subset to the batch analyzer.  Keeping the implementation
// here makes lint, LSP, analyze, and preflight share exactly the same symbol
// and constant-bound decisions without a second rule-specific inference path.
func CompileEquivalentArrayShapeIssues(path, source string, ir *procedureir.DocumentIR) []Issue {
	return CompileEquivalentArrayShapeIssuesWithConstants(path, source, ir, nil)
}

// CompileEquivalentArrayShapeIssuesWithConstants is the project-aware form
// used by batch/LSP adapters that already have a coherent visible-constant
// table. Keeping the optional table at the adapter boundary avoids a second
// resolver or expression-inference implementation in this rule.
func CompileEquivalentArrayShapeIssuesWithConstants(path, source string, ir *procedureir.DocumentIR, visibleConstants map[string]bool) []Issue {
	return (Linter{VisibleConstants: visibleConstants}).arrayShapeIssues(path, source, ir)
}

// CompileEquivalentArrayShapeIssuesWithValues is the value-bearing project
// adapter used by batch/LSP callers that already own a coherent snapshot.
func CompileEquivalentArrayShapeIssuesWithValues(path, source string, ir *procedureir.DocumentIR, visibleConstants map[string]bool, values map[string]constexpr.Value) []Issue {
	return (Linter{VisibleConstants: visibleConstants, ConstantValues: values}).arrayShapeIssues(path, source, ir)
}

func declarationPrefix(lower string) bool {
	for _, kw := range []string{"dim ", "public ", "private ", "friend ", "static "} {
		if strings.HasPrefix(lower, kw) {
			return true
		}
	}
	return false
}

func declarationBody(line string) string {
	fields := strings.Fields(line)
	if len(fields) == 0 {
		return ""
	}
	idx := strings.Index(strings.ToLower(line), strings.ToLower(fields[0]))
	for idx < len(line) && !unicode.IsSpace(rune(line[idx])) {
		idx++
	}
	return strings.TrimSpace(line[idx:])
}

func splitArrayDeclarators(text string) []string {
	var out []string
	start, depth := 0, 0
	for i, r := range text {
		switch r {
		case '(':
			depth++
		case ')':
			if depth > 0 {
				depth--
			}
		case ',':
			if depth == 0 {
				out = append(out, strings.TrimSpace(text[start:i]))
				start = i + 1
			}
		}
	}
	if start < len(text) {
		out = append(out, strings.TrimSpace(text[start:]))
	}
	return out
}

func arrayDeclaratorBounds(part string) (string, string, bool) {
	open := strings.Index(part, "(")
	if open <= 0 {
		return "", "", false
	}
	name := strings.TrimSpace(part[:open])
	if !isArrayDeclaratorName(name) {
		return "", "", false
	}
	depth, close := 0, -1
	for i := open; i < len(part); i++ {
		if part[i] == '(' {
			depth++
		}
		if part[i] == ')' {
			depth--
			if depth == 0 {
				close = i
				break
			}
		}
	}
	if close < 0 {
		return "", "", false
	}
	return name, part[open+1 : close], true
}

func isArrayDeclaratorName(text string) bool {
	text = strings.TrimSpace(text)
	if text == "" {
		return false
	}
	last := text[len(text)-1]
	if strings.ContainsRune("$%&!#@^", rune(last)) {
		text = text[:len(text)-1]
	}
	if text == "" || (text[0] != '_' && !unicode.IsLetter(rune(text[0]))) {
		return false
	}
	for _, r := range text[1:] {
		if r != '_' && !unicode.IsLetter(r) && !unicode.IsDigit(r) {
			return false
		}
	}
	return true
}

func evalArrayConstant(expr string, consts map[string]bool, values map[string]int) (int, bool) {
	if values == nil {
		values = map[string]int{}
	}
	return parseArrayConstant(strings.TrimSpace(expr), consts, values)
}

func parseArrayConstant(expr string, consts map[string]bool, values map[string]int) (int, bool) {
	_ = consts // names are represented by the values map once resolved
	result := constexpr.EvaluateInteger(strings.TrimSpace(expr), values)
	return result.Value, result.Kind == constexpr.Known
}
