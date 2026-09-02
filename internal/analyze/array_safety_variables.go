package analyze

import (
	"strconv"
	"strings"

	"github.com/harumiWeb/xlflow/internal/gui"
	"github.com/harumiWeb/xlflow/internal/lint"
	"github.com/harumiWeb/xlflow/internal/vba/constexpr"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func arrayVariables(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) map[string]arrayVariable {
	decls := newDeclarationScope(file, proc)
	decls.module = moduleDecls
	base := arrayOptionBase(file)
	// The legacy line scanner intentionally ignores Const and Enum syntax.
	// Fill those gaps from the shared procedure IR so a known scalar constant
	// can still be classified as a non-iterable/non-array operation target.
	for declaration := range proc.Declarations.All() {
		key := strings.ToLower(declaration.Name)
		if key == "" {
			continue
		}
		decls.addExtraIfMissing(key, sourceDeclaration{
			Name: declaration.Name, Type: declaration.Type, Line: declaration.Range.StartLine,
			Object: declaration.IsObject, Array: declaration.IsArray,
			Fixed:      declaration.ValueShape == procedureir.ValueShapeFixedArray,
			Dimensions: parameterArrayDimensions(declaration.ArrayBounds, base),
		})
	}
	for param := range proc.Params.All() {
		array := param.ParamArray || strings.Contains(param.Type, "()") || param.ValueShape == procedureir.ValueShapeFixedArray || param.ValueShape == procedureir.ValueShapeDynamicArray
		decls.parameters[strings.ToLower(param.Name)] = sourceDeclaration{
			Name: param.Name, Type: param.Type, Array: array,
			Fixed:      param.ValueShape == procedureir.ValueShapeFixedArray,
			Dimensions: parameterArrayDimensions(param.ArrayBounds, base),
			Object:     isObjectType(param.Type), Parameter: true, ParamArray: param.ParamArray,
		}
	}
	catalog := file.ArrayVariableCatalog
	if catalog == nil {
		catalog = buildArrayVariableCatalog(file, moduleDecls)
	}
	usedModuleNames := make(map[string]struct{}, proc.Accesses.Len())
	for access := range proc.Accesses.All() {
		if name := strings.ToLower(strings.TrimSpace(access.Name)); name != "" {
			usedModuleNames[name] = struct{}{}
		}
	}
	// An empty access projection can mean a genuinely independent procedure,
	// but it can also be the only reliable view left after parser recovery or
	// an incomplete IR build. Preserve the historical fail-open behavior for
	// that boundary instead of silently dropping module scalars from array
	// classification.
	includeAllModule := proc.Accesses.Len() == 0
	moduleCapacity := len(usedModuleNames)
	for _, variable := range catalog {
		if includeAllModule || variable.isArray {
			moduleCapacity++
		}
	}
	variables := make(map[string]arrayVariable, moduleCapacity+len(decls.extra)+len(decls.local)+len(decls.parameters))
	for key, variable := range catalog {
		if !includeAllModule && !variable.isArray {
			if _, used := usedModuleNames[key]; !used {
				continue
			}
		}
		variables[key] = variable
	}
	// Module declarations are immutable and already normalized in catalog.
	// Only the procedure overlays need to be materialized here. Overlay order
	// matches declarationScope.forEach: IR extras, source locals, parameters.
	for key, decl := range decls.extra {
		variables[key] = newArrayVariable(file, decl, base)
	}
	for key, decl := range decls.local {
		variables[key] = newArrayVariable(file, decl, base)
	}
	for key, decl := range decls.parameters {
		variables[key] = newArrayVariable(file, decl, base)
	}
	return variables
}

func buildArrayVariableCatalog(file parsedFile, moduleDecls map[string]sourceDeclaration) map[string]arrayVariable {
	variables := make(map[string]arrayVariable, len(moduleDecls))
	base := arrayOptionBase(file)
	for key, decl := range moduleDecls {
		variables[key] = newArrayVariable(file, decl, base)
	}
	return variables
}

func newArrayVariable(file parsedFile, decl sourceDeclaration, base int) arrayVariable {
	typeName := strings.TrimSpace(decl.Type)
	// VBA declarations without an As clause are implicit Variant values.
	// Treat them exactly like an explicit Variant so array-sensitive rules
	// do not turn an unresolved value into a false scalar proof.
	isVariant := typeName == "" || strings.EqualFold(typeName, "Variant")
	variable := arrayVariable{name: decl.Name, typ: decl.Type, isArray: decl.Array, isVariant: isVariant, isObject: decl.Object, knownScalar: !isVariant && arrayKnownScalarType(typeName), parameter: decl.Parameter, static: decl.Static, paramArray: decl.ParamArray}
	if decl.Array {
		variable.dimensions, variable.fixed = declarationDimensions(file.Lines, decl.Line, decl.Name, base)
		if len(decl.Dimensions) > 0 {
			variable.dimensions = append([]arrayDimension(nil), decl.Dimensions...)
		}
		if decl.Fixed {
			variable.fixed = true
		}
		if len(variable.dimensions) > 0 {
			variable.fixed = true
		}
	}
	return variable
}

func isByteArrayVariable(variable arrayVariable) bool {
	return variable.isArray && isByteArrayTypeName(variable.typ)
}

func isByteArrayTypeName(typeName string) bool {
	typeName = strings.TrimSpace(typeName)
	if colon := strings.IndexByte(typeName, ':'); colon >= 0 {
		typeName = strings.TrimSpace(typeName[:colon])
	}
	return strings.EqualFold(typeName, "Byte")
}

// arrayVBA227Variables overlays the narrow declaration facts needed by the
// source-line VBA227 pass. The legacy declaration scanner is intentionally
// retained for the historical array rules, but it can span a colon-separated
// `Dim ...: ReDim ...` line and mistake the ReDim bounds for declaration
// bounds. Procedure IR identifies the declaration's dynamic shape without
// that ambiguity.
func arrayVBA227Variables(baseVariables map[string]arrayVariable, file parsedFile, proc sourceProcedure) map[string]arrayVariable {
	overlays := make([]procedureir.Declaration, 0)
	base := arrayOptionBase(file)
	for declaration := range proc.Declarations.All() {
		key := strings.ToLower(strings.TrimSpace(declaration.Name))
		if key == "" || !declaration.IsArray || declaration.ValueShape != procedureir.ValueShapeDynamicArray {
			continue
		}
		inlineRedim := arrayDeclarationHasInlineRedim(file.Lines, declaration)
		inlineStrConv := isByteArrayTypeName(declaration.Type) && arrayDeclarationHasInlineStrConvAssignment(file.Lines, declaration)
		if !inlineRedim && !inlineStrConv {
			continue
		}
		overlays = append(overlays, declaration)
	}
	if len(overlays) == 0 {
		return baseVariables
	}
	variables := make(map[string]arrayVariable, len(baseVariables))
	for key, variable := range baseVariables {
		variable.dimensions = append([]arrayDimension(nil), variable.dimensions...)
		variables[key] = variable
	}
	for _, declaration := range overlays {
		key := strings.ToLower(strings.TrimSpace(declaration.Name))
		variable, ok := variables[key]
		if !ok {
			variable = arrayVariable{
				name:        declaration.Name,
				typ:         declaration.Type,
				isArray:     true,
				isVariant:   strings.EqualFold(strings.TrimSpace(declaration.Type), "Variant"),
				isObject:    declaration.IsObject,
				knownScalar: false,
			}
		}
		variable.name = declaration.Name
		variable.typ = declaration.Type
		variable.isArray = true
		variable.isVariant = strings.EqualFold(strings.TrimSpace(declaration.Type), "Variant")
		variable.isObject = declaration.IsObject
		variable.dimensions = parameterArrayDimensions(declaration.ArrayBounds, base)
		variable.fixed = declaration.ValueShape == procedureir.ValueShapeFixedArray || len(variable.dimensions) > 0
		variables[key] = variable
	}
	return variables
}

func arrayDeclarationHasInlineRedim(lines []string, declaration procedureir.Declaration) bool {
	line := declaration.Range.StartLine
	if line < 1 || line > len(lines) {
		return false
	}
	text := strings.ToLower(normalizedCodeLine(lines[line-1]))
	colon := strings.IndexByte(text, ':')
	if colon < 0 {
		return false
	}
	return strings.HasPrefix(strings.TrimSpace(text[colon+1:]), "redim ")
}

func arrayDeclarationHasInlineStrConvAssignment(lines []string, declaration procedureir.Declaration) bool {
	line := declaration.Range.StartLine
	if line < 1 || line > len(lines) {
		return false
	}
	parts := splitRangeValueSourceStatements(normalizedCodeLine(lines[line-1]))
	for _, part := range parts[1:] {
		lhs, rhs, indexed, ok := arrayAssignment(part)
		if ok && !indexed && strings.EqualFold(lhs, declaration.Name) && strings.EqualFold(arrayCallName(rhs), "strconv") {
			return true
		}
	}
	return false
}

func parameterArrayDimensions(bounds []procedureir.ArrayBound, base int) []arrayDimension {
	if len(bounds) == 0 {
		return nil
	}
	dimensions := make([]arrayDimension, 0, len(bounds))
	for _, bound := range bounds {
		lowerText, upperText := strings.TrimSpace(bound.Lower), strings.TrimSpace(bound.Upper)
		if lowerText == "" && upperText == "" {
			upperText = strings.TrimSpace(bound.Expression)
			lowerText = strconv.Itoa(base)
		}
		dimensions = append(dimensions, arrayDimension{
			lower: integerBound(lowerText), upper: integerBound(upperText),
		})
	}
	return dimensions
}

func optionBase(lines []string) int {
	for _, line := range lines {
		fields := strings.Fields(strings.ToLower(normalizedCodeLine(line)))
		if len(fields) == 3 && fields[0] == "option" && fields[1] == "base" {
			if value, err := strconv.Atoi(fields[2]); err == nil && (value == 0 || value == 1) {
				return value
			}
		}
	}
	return 0
}

func arrayOptionBase(file parsedFile) int {
	if file.ArrayOptionBaseSet {
		return file.ArrayOptionBase
	}
	return optionBase(file.Lines)
}

func declarationDimensions(lines []string, line int, name string, base int) ([]arrayDimension, bool) {
	if line < 1 || line > len(lines) {
		return nil, false
	}
	stmt := normalizedCodeLine(lines[line-1])
	m := declRe.FindStringSubmatch(stmt)
	if len(m) == 0 {
		return nil, false
	}
	for _, part := range splitArgs(m[1]) {
		candidate, _, array, _ := declarationNameAndType(part)
		if !array || !strings.EqualFold(candidate, name) {
			continue
		}
		start := strings.Index(part, "(")
		end := strings.LastIndex(part, ")")
		if start < 0 || end < start {
			return nil, false
		}
		raw := strings.TrimSpace(part[start+1 : end])
		if raw == "" {
			return nil, false
		}
		return parseArrayDimensions(raw, base), true
	}
	return nil, false
}

func parseArrayDimensions(text string, base int) []arrayDimension {
	var dimensions []arrayDimension
	for _, part := range splitArgs(text) {
		part = strings.TrimSpace(part)
		if part == "" {
			continue
		}
		pieces := strings.SplitN(strings.ToLower(part), " to ", 2)
		if len(pieces) == 2 {
			dimensions = append(dimensions, arrayDimension{lower: integerBound(pieces[0]), upper: integerBound(pieces[1])})
			continue
		}
		upper := integerBound(part)
		dimensions = append(dimensions, arrayDimension{lower: integerBound(strconv.Itoa(base)), upper: upper})
	}
	return dimensions
}

func parseArrayDimensionsWithConstants(text string, base int, constants map[string]int) []arrayDimension {
	var dimensions []arrayDimension
	for _, part := range splitArgs(text) {
		part = strings.TrimSpace(part)
		if part == "" {
			continue
		}
		pieces := strings.SplitN(strings.ToLower(part), " to ", 2)
		if len(pieces) == 2 {
			dimensions = append(dimensions, arrayDimension{lower: integerBoundWithConstants(pieces[0], constants), upper: integerBoundWithConstants(pieces[1], constants)})
			continue
		}
		dimensions = append(dimensions, arrayDimension{lower: integerBoundWithConstants(strconv.Itoa(base), constants), upper: integerBoundWithConstants(part, constants)})
	}
	return dimensions
}

// arrayIntegerConstants extends the shared range-value constant table with
// the Enum members that are valid integer bound names in VBA.  The parser is
// intentionally small and fail-open: unresolved expressions simply do not
// enter the table, so dynamic bounds remain unclassified.
func arrayIntegerConstants(file parsedFile, proc sourceProcedure, projectValues map[string]constexpr.Value, visibleNames map[string]bool) map[string]int {
	constants := rangeValueIntegerConstants(arrayIntegerModuleConstants(file), proc)
	// Project constants fill names that are not already provided by the module
	// projection, preserving the historical module-over-project precedence.
	for name, value := range projectValues {
		visibleKey := strings.ToLower(cleanIdentifier(name))
		if visibleNames != nil && !visibleNames[visibleKey] {
			continue
		}
		if value.Kind != constexpr.ValueInteger && value.Kind != constexpr.ValueLong && value.Kind != constexpr.ValueLongLong {
			continue
		}
		if integer, ok := constexpr.IntegerAsInt(value); ok {
			key := strings.ToLower(name)
			if _, exists := constants[key]; !exists {
				constants[key] = integer
			}
		}
	}
	return constants
}

func arrayIntegerModuleConstants(file parsedFile) map[string]int {
	if file.ArrayIntegerModuleConstants != nil {
		return file.ArrayIntegerModuleConstants
	}
	base := file.RangeValueModuleConstants
	if base == nil {
		base = rangeValueModuleIntegerConstants(file.Lines, file.IR)
	}
	constants := make(map[string]int, len(base))
	for name, value := range base {
		constants[name] = value
	}
	sharedValues := file.ConstantValues
	if sharedValues == nil {
		sharedValues = lint.ConstantValuesFromSource(string(file.Source), &file.IR, nil)
	}
	for name, value := range sharedValues {
		if value.Kind != constexpr.ValueInteger && value.Kind != constexpr.ValueLong && value.Kind != constexpr.ValueLongLong {
			continue
		}
		if integer, ok := constexpr.IntegerAsInt(value); ok {
			constants[strings.ToLower(name)] = integer
		}
	}
	// Keep the legacy enum fallback for hand-built parsedFile values whose
	// ConstantValues projection is absent or incomplete.
	inEnum := false
	var next *int
	conditionalDepth := 0
	for _, line := range file.Lines {
		code := strings.TrimSpace(normalizedCodeLine(line))
		lower := strings.ToLower(code)
		if strings.HasPrefix(lower, "#if ") {
			conditionalDepth++
			continue
		}
		if strings.HasPrefix(lower, "#end if") {
			if conditionalDepth > 0 {
				conditionalDepth--
			}
			continue
		}
		if conditionalDepth > 0 || strings.HasPrefix(lower, "#elseif ") || strings.HasPrefix(lower, "#else") {
			continue
		}
		if strings.HasPrefix(lower, "enum ") || strings.HasPrefix(lower, "public enum ") || strings.HasPrefix(lower, "private enum ") || strings.HasPrefix(lower, "friend enum ") {
			inEnum = true
			next = nil
			continue
		}
		if inEnum && strings.HasPrefix(lower, "end enum") {
			inEnum = false
			next = nil
			continue
		}
		if !inEnum || code == "" {
			continue
		}
		parts := strings.SplitN(code, "=", 2)
		fields := strings.Fields(parts[0])
		if len(fields) == 0 {
			continue
		}
		name := strings.TrimSpace(fields[0])
		if name == "" {
			continue
		}
		key := strings.ToLower(cleanIdentifier(name))
		if len(parts) == 2 {
			value, err := constantIntegerExpression(strings.TrimSpace(parts[1]), constants)
			if err != nil {
				next = nil
				continue
			}
			constants[key] = value
			n := value + 1
			next = &n
			continue
		}
		if next == nil {
			continue
		}
		constants[key] = *next
		n := *next + 1
		next = &n
	}
	return constants
}

func integerBound(text string) arrayBound {
	value, ok := integerLiteral(text)
	return arrayBound{known: ok, value: value, expression: canonicalArrayBoundExpression(text)}
}

func integerBoundWithConstants(text string, constants map[string]int) arrayBound {
	value, err := constantIntegerExpression(text, constants)
	if err != nil {
		return arrayBound{expression: canonicalArrayBoundExpression(text)}
	}
	return arrayBound{known: true, value: value, expression: canonicalArrayBoundExpression(text)}
}

func impossibleArrayBounds(dimensions []arrayDimension) bool {
	for _, dimension := range dimensions {
		if dimension.lower.known && dimension.upper.known && dimension.lower.value > dimension.upper.value {
			return true
		}
	}
	return false
}

func iterableSourceKnownInvalid(source string, variables map[string]arrayVariable, state arrayFlowState, ctx analysisContext) bool {
	source = strings.TrimSpace(source)
	if source == "" {
		return false
	}
	// A uniquely resolved scalar Function/Property Get return is as strong as
	// a scalar declaration.  Unknown, external, or duplicate call targets are
	// deliberately absent from functionShapes and therefore remain fail-open.
	if strings.Contains(source, "(") {
		callee := arrayCallName(source)
		if returnType, ok := ctx.functionReturns[callee]; ok && isObjectType(returnType) {
			// Known Collection/Dictionary/Object-compatible returns are valid
			// iterable sources even though their value shape is non-array.
			return false
		}
		if shape, ok := ctx.functionShapes[callee]; ok {
			switch shape {
			case procedureir.ValueShapeScalar:
				return true
			case procedureir.ValueShapeFixedArray, procedureir.ValueShapeDynamicArray:
				return false
			}
		}
	}
	// An indexed element of a statically typed scalar array is itself a
	// scalar, even though the base expression is an allocated array. Variant
	// and object arrays remain conservative because their elements may hold an
	// array or implement the Collection iteration contract.
	if match := arrayIndexedSourceRe.FindStringSubmatch(source); len(match) > 0 {
		if variable, ok := variables[strings.ToLower(match[1])]; ok && variable.isArray && !variable.isVariant && !variable.isObject && variable.knownScalar {
			return true
		}
	}
	if value, known := arrayExpressionState(source, state, ctx); known {
		if value.knownArray {
			return false
		}
		// Unknown calls and Variant values remain conservative.  A known scalar
		// expression is handled below where its declaration is available.
		if strings.Contains(source, "(") {
			return false
		}
	}
	if scalarExpressionKnown(source) {
		return true
	}
	name := source
	if dot := strings.IndexAny(name, ".("); dot >= 0 {
		name = name[:dot]
	}
	name = strings.TrimSpace(name)
	if !arrayEraseNameRe.MatchString(name) {
		// Numeric, Boolean, date, and string literals are definitely scalar; all other
		// expressions remain unknown rather than guessing their result type.
		return false
	}
	variable, ok := variables[strings.ToLower(name)]
	if !ok || variable.isArray || variable.isVariant || variable.isObject || !variable.knownScalar || dcKindFromType(variable.typ) != dcUnknown {
		return false
	}
	return true
}

func arrayKnownScalarType(typ string) bool {
	switch strings.ToLower(cleanIdentifier(strings.TrimSpace(typ))) {
	case "byte", "boolean", "integer", "long", "longlong", "longptr", "single", "double", "currency", "decimal", "date", "string":
		return true
	default:
		return false
	}
}

func scalarExpressionKnown(text string) bool {
	text = strings.TrimSpace(text)
	if text == "" {
		return false
	}
	lower := strings.ToLower(text)
	switch lower {
	case "true", "false", "nothing", "empty", "null":
		return true
	}
	if strings.HasPrefix(text, `"`) || (strings.HasPrefix(text, "#") && strings.HasSuffix(text, "#")) {
		return true
	}
	if _, ok := integerLiteral(text); ok {
		return true
	}
	if _, err := strconv.ParseFloat(strings.TrimRight(text, "%&!@$^"), 64); err == nil {
		return true
	}
	_, err := constantIntegerExpression(text, nil)
	return err == nil
}

func integerLiteral(text string) (int, bool) {
	value, err := strconv.Atoi(strings.TrimSpace(text))
	return value, err == nil
}

func preserveDimensionsSafe(previous, next []arrayDimension) bool {
	if len(next) == 1 {
		return len(previous) <= 1
	}
	if len(previous) != len(next) || len(previous) == 0 {
		return false
	}
	for i := 0; i < len(previous)-1; i++ {
		if !arrayBoundsEquivalent(previous[i].lower, next[i].lower) || !arrayBoundsEquivalent(previous[i].upper, next[i].upper) {
			return false
		}
	}
	return true
}

func arrayBoundsEquivalent(left, right arrayBound) bool {
	if left.known && right.known {
		return left.value == right.value
	}
	return left.expression != "" && left.expression == right.expression
}

func canonicalArrayBoundExpression(text string) string {
	var out strings.Builder
	inString := false
	for i := 0; i < len(text); i++ {
		ch := text[i]
		if ch == '"' {
			out.WriteByte(ch)
			if inString && i+1 < len(text) && text[i+1] == '"' {
				out.WriteByte(text[i+1])
				i++
				continue
			}
			inString = !inString
			continue
		}
		if !inString {
			if ch == ' ' || ch == '\t' || ch == '\r' || ch == '\n' {
				continue
			}
			if ch >= 'A' && ch <= 'Z' {
				ch += 'a' - 'A'
			}
		}
		out.WriteByte(ch)
	}
	return out.String()
}

func parseDirectArrayRedimClause(clause string) (directArrayRedimClause, bool) {
	clause = strings.TrimSpace(clause)
	if clause == "" || !isIdentifierStart(clause[0]) {
		return directArrayRedimClause{}, false
	}
	endName := 1
	for endName < len(clause) && isIdentifierPart(clause[endName]) {
		endName++
	}
	open := endName
	for open < len(clause) && (clause[open] == ' ' || clause[open] == '\t') {
		open++
	}
	if open >= len(clause) || clause[open] != '(' {
		return directArrayRedimClause{}, false
	}
	close := matchingParen(clause, open)
	if close < 0 {
		return directArrayRedimClause{}, false
	}
	remainder := strings.TrimSpace(clause[close+1:])
	if remainder != "" && !arrayRedimTypeSuffixRe.MatchString(remainder) {
		return directArrayRedimClause{}, false
	}
	dimensions := strings.TrimSpace(clause[open+1 : close])
	if dimensions == "" {
		return directArrayRedimClause{}, false
	}
	return directArrayRedimClause{name: clause[:endName], dimensions: dimensions}, true
}

func arrayIndexedUses(text string, variables map[string]arrayVariable) []arrayUse {
	var uses []arrayUse
	for i := 0; i < len(text); i++ {
		if !isIdentifierStart(text[i]) || (i > 0 && isIdentifierPart(text[i-1])) {
			continue
		}
		start := i
		i++
		for i < len(text) && isIdentifierPart(text[i]) {
			i++
		}
		name := text[start:i]
		// A qualified member call such as Application.OnTime(...) can
		// coincide with a local scalar named OnTime.  The member name is not
		// an array variable access; only unqualified identifiers participate
		// in this source-level array scan.
		if start > 0 && (text[start-1] == '.' || text[start-1] == '!') {
			continue
		}
		key := strings.ToLower(name)
		variable, ok := variables[key]
		if !ok || !variable.isArray && !variable.isVariant {
			continue
		}
		for i < len(text) && (text[i] == ' ' || text[i] == '\t') {
			i++
		}
		if i >= len(text) || text[i] != '(' {
			continue
		}
		end := matchingParen(text, i)
		if end < 0 {
			continue
		}
		uses = append(uses, arrayUse{name: name, args: splitArgs(text[i+1 : end])})
		i = end
	}
	return uses
}

// arrayIndexedUsesForSource scans source-line CFG text after removing syntax
// that cannot perform an array access. Recovered blocks may retain string
// literals and comments, so feeding their raw text to the VBA227 lane would
// mistake prose such as "clipboard data (" for an indexed use. Keep the
// primitive scanner raw for runtime projections that have their own source
// normalization contract.
func arrayIndexedUsesForSource(text string, variables map[string]arrayVariable) []arrayUse {
	return arrayIndexedUses(maskStringLiterals(gui.StripComment(text)), variables)
}

func matchingParen(text string, start int) int {
	depth := 0
	inString := false
	for i := start; i < len(text); i++ {
		switch text[i] {
		case '"':
			if inString && i+1 < len(text) && text[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
		case '(':
			if !inString {
				depth++
			}
		case ')':
			if !inString {
				depth--
				if depth == 0 {
					return i
				}
			}
		}
	}
	return -1
}

func isIdentifierStart(char byte) bool {
	return char == '_' || char >= 'A' && char <= 'Z' || char >= 'a' && char <= 'z'
}
func isIdentifierPart(char byte) bool { return isIdentifierStart(char) || char >= '0' && char <= '9' }

func arrayAssignment(text string) (lhs, rhs string, indexed, ok bool) {
	trimmed := strings.TrimSpace(text)
	if strings.HasPrefix(strings.ToLower(trimmed), "set ") {
		trimmed = strings.TrimSpace(trimmed[4:])
	}
	if strings.HasPrefix(strings.ToLower(trimmed), "let ") {
		trimmed = strings.TrimSpace(trimmed[4:])
	}
	for i := 0; i < len(trimmed); i++ {
		if trimmed[i] != '=' || i > 0 && (trimmed[i-1] == '<' || trimmed[i-1] == '>' || trimmed[i-1] == '=') {
			continue
		}
		lhs = strings.TrimSpace(trimmed[:i])
		rhs = strings.TrimSpace(trimmed[i+1:])
		if lhs == "" || strings.HasPrefix(strings.ToLower(lhs), "if ") {
			return "", "", false, false
		}
		if open := strings.Index(lhs, "("); open >= 0 && strings.HasSuffix(lhs, ")") {
			wholeArray := strings.TrimSpace(lhs[open+1:len(lhs)-1]) == ""
			return cleanIdentifier(lhs[:open]), rhs, !wholeArray, true
		}
		return cleanIdentifier(lhs), rhs, false, true
	}
	return "", "", false, false
}

func arrayExpressionState(rhs string, state arrayFlowState, ctx analysisContext) (arrayValue, bool) {
	lower := strings.ToLower(strings.TrimSpace(rhs))
	if strings.HasPrefix(lower, "array(") || lower == "array" {
		shape := []arrayDimension{{}}
		return arrayValue{kind: arrayAllocated, knownArray: true, dimensions: shape, preserveShape: shape, origin: arrayOriginLocal}, true
	}
	if strings.Contains(lower, ".value") || strings.Contains(lower, ".value2") {
		return arrayValue{kind: arrayAllocated, knownArray: true, origin: arrayOriginRangeValue}, true
	}
	name := arrayCallName(rhs)
	if name == "split" || name == "filter" {
		shape := []arrayDimension{{}}
		return arrayValue{kind: arrayAllocated, knownArray: true, dimensions: shape, preserveShape: shape, origin: arrayOriginLocal}, true
	}
	// Keep StrConv-to-Byte-array assignments in the type-aware transfer below;
	// it can distinguish a known non-empty String from an unknown one.
	if name == "strconv" {
		return arrayValue{}, false
	}
	if arrayByteArrayReadRe.MatchString(rhs) {
		return arrayValue{kind: arrayAllocated, knownArray: true, origin: arrayOriginLocal}, true
	}
	if value, ok := state[name]; ok && value.kind == arrayAllocated && value.knownArray {
		return value, true
	}
	if value, ok := ctx.arrayReturns[name]; ok {
		_, hasLowerBound := arrayValueKnownLowerBound(value)
		if value.knownArray || hasLowerBound {
			return value, true
		}
	}
	if name != "" && strings.Contains(rhs, "(") {
		return arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}, true
	}
	if strings.EqualFold(strings.TrimSpace(rhs), "empty") || strings.EqualFold(strings.TrimSpace(rhs), "nothing") {
		return arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}, true
	}
	return arrayValue{}, false
}
