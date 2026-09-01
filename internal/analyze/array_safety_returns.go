package analyze

import (
	"regexp"
	"sort"
	"strings"

	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// arrayVBA227AttachReturnProvenance consumes the narrow conditional facts in
// an array-return summary at the call site. The summary records a formal
// parameter; this step maps it to the actual argument and keeps the fact
// path-sensitive until the caller proves the corresponding condition.
func arrayVBA227AttachReturnProvenance(state arrayFlowState, text string, ctx analysisContext, variables map[string]arrayVariable, constants map[string]int) arrayFlowState {
	lhs, rhs, indexed, ok := arrayAssignment(text)
	if !ok || indexed {
		return state
	}
	callee := arrayCallName(rhs)
	if callee == "" || ctx.functionAmbiguous[callee] {
		return state
	}
	summary, ok := ctx.arrayReturns[callee]
	if !ok || summary.returnNonEmptyArrayParameter == "" && summary.returnPositiveScalarParameter == "" && summary.returnDescriptorSourceParameter == "" {
		return state
	}
	target := strings.ToLower(cleanIdentifier(lhs))
	variable, knownVariable := variables[target]
	value, knownValue := state[target]
	if !knownVariable || !knownValue || !variable.isArray && !variable.isVariant {
		return state
	}
	signature, ok := ctx.procedures[callee]
	if !ok {
		return state
	}
	arguments, ok := arrayReturnCallArguments(rhs, signature)
	if !ok {
		return state
	}
	updated := value
	changed := false
	if formal := summary.returnPositiveScalarParameter; formal != "" {
		if index, found := arrayFormalParameterIndexFromSignature(signature, formal); found && index < len(arguments) {
			if length, err := constantIntegerExpression(strings.TrimSpace(arguments[index]), constants); err == nil && length > 0 {
				updated.kind = arrayAllocated
				updated.knownArray = true
				updated.mayBeEmpty = false
				updated.returnPositiveScalarParameter = ""
				updated.returnNonEmptyArrayParameter = ""
				changed = true
			}
		}
	}
	if formal := summary.returnNonEmptyArrayParameter; formal != "" {
		if index, found := arrayFormalParameterIndexFromSignature(signature, formal); found && index < len(arguments) {
			if source := directArrayArgumentName(arguments[index]); source != "" {
				sourceVariable, sourceKnown := variables[source]
				if sourceKnown && sourceVariable.isArray {
					updated.nonEmptySource = source
					updated.returnNonEmptyArrayParameter = ""
					updated.returnPositiveScalarParameter = ""
					changed = true
				}
			}
		}
	}
	if summary.returnDescriptorSourceParameter != "" {
		updated = arrayDescriptorReturnCallValue(updated, summary, arguments, signature, constants)
		updated.returnDescriptorSourceParameter = ""
		updated.returnDescriptorStartParameter = ""
		updated.returnDescriptorLengthParameter = ""
		updated.returnDescriptorLowerParameter = ""
		changed = true
	}
	if !changed {
		return state
	}
	updatedState := cloneArrayState(state)
	updatedState[target] = updated
	return updatedState
}

func arraySimpleCallArguments(text string) ([]string, bool) {
	text = strings.TrimSpace(text)
	open := firstParenOutsideString(text)
	if open < 0 {
		return nil, false
	}
	close := matchingParen(text, open)
	if close < 0 || strings.TrimSpace(text[close+1:]) != "" {
		return nil, false
	}
	return splitArgs(text[open+1 : close]), true
}

var arrayNamedArgumentRe = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*:=\s*(.*?)\s*$`)

// arrayReturnCallArguments maps positional and named actuals to a procedure
// signature and fills omitted optional arguments with their declared defaults.
// Return summaries need this small binding layer because a call such as
// StringToIntegers("ABC", outLowBound:=5) cannot be interpreted by position.
func arrayReturnCallArguments(text string, signature procedureSignature) ([]string, bool) {
	raw, ok := arraySimpleCallArguments(text)
	if !ok {
		return nil, false
	}
	arguments := make([]string, signature.Params.Len())
	assigned := make([]bool, len(arguments))
	nextPositional := 0
	for _, rawArgument := range raw {
		argument := strings.TrimSpace(rawArgument)
		if match := arrayNamedArgumentRe.FindStringSubmatch(argument); len(match) == 3 {
			index, found := arrayFormalParameterIndexFromSignature(signature, match[1])
			if !found || assigned[index] {
				return nil, false
			}
			arguments[index] = strings.TrimSpace(match[2])
			assigned[index] = true
			continue
		}
		for nextPositional < len(assigned) && assigned[nextPositional] {
			nextPositional++
		}
		if nextPositional >= len(arguments) {
			return nil, false
		}
		if argument == "" {
			nextPositional++
			continue
		}
		arguments[nextPositional] = argument
		assigned[nextPositional] = true
		nextPositional++
	}
	for index, parameter := range signature.Params.AllIndexed() {
		if assigned[index] {
			continue
		}
		if !parameter.Optional {
			return nil, false
		}
		if parameter.HasDefault {
			arguments[index] = parameter.Default
		}
	}
	return arguments, true
}

func arrayDescriptorReturnCallValue(value, summary arrayValue, arguments []string, signature procedureSignature, constants map[string]int) arrayValue {
	argument := func(formal string) (string, bool) {
		index, ok := arrayFormalParameterIndexFromSignature(signature, formal)
		if !ok || index >= len(arguments) {
			return "", false
		}
		return strings.TrimSpace(arguments[index]), true
	}
	source, sourceOK := argument(summary.returnDescriptorSourceParameter)
	startText, startOK := argument(summary.returnDescriptorStartParameter)
	lengthText, lengthOK := argument(summary.returnDescriptorLengthParameter)
	lowerText, lowerOK := argument(summary.returnDescriptorLowerParameter)
	if !sourceOK || !startOK || !lengthOK || !lowerOK {
		return value
	}
	start, startErr := constantIntegerExpression(startText, constants)
	length, lengthErr := constantIntegerExpression(lengthText, constants)
	lower, lowerErr := constantIntegerExpression(lowerText, constants)
	if startErr != nil || lengthErr != nil || lowerErr != nil || start < 1 || length < -1 {
		return value
	}

	count, countKnown := 0, false
	if length == 0 {
		count, countKnown = 0, true
	} else if sourceLength, known := arrayStringExpressionKnownLength(source); known {
		count = length
		if length == -1 || start+length-1 > sourceLength {
			count = sourceLength - start + 1
			if count < 0 {
				count = 0
			}
		}
		countKnown = true
	}
	dimension := arrayDimension{lower: arrayBound{known: true, value: lower}}
	if countKnown {
		dimension.upper = arrayBound{known: true, value: lower + count - 1}
		value.mayBeEmpty = count == 0
	} else {
		// A descriptor-backed return always has a valid SAFEARRAY descriptor,
		// but an unknown input length may still produce zero elements.
		value.mayBeEmpty = true
	}
	value.dimensions = []arrayDimension{dimension}
	value.preserveShape = append([]arrayDimension(nil), value.dimensions...)
	return value
}

func arrayStringExpressionKnownLength(expression string) (int, bool) {
	expression = strings.TrimSpace(expression)
	if strings.EqualFold(expression, "vbNullString") {
		return 0, true
	}
	if length, ok := arrayStringLiteralLength(expression); ok {
		return length, true
	}
	if !strings.EqualFold(arrayCallName(expression), "strconv") {
		return 0, false
	}
	open := firstParenOutsideString(expression)
	close := matchingParen(expression, open)
	if open < 0 || close < 0 || strings.TrimSpace(expression[close+1:]) != "" {
		return 0, false
	}
	arguments := splitArgs(expression[open+1 : close])
	if len(arguments) == 0 {
		return 0, false
	}
	return arrayStringExpressionKnownLength(arguments[0])
}

func arrayStringLiteralLength(expression string) (int, bool) {
	if len(expression) < 2 || expression[0] != '"' || expression[len(expression)-1] != '"' {
		return 0, false
	}
	length := 0
	for index := 1; index < len(expression)-1; index++ {
		if expression[index] >= 0x80 {
			return 0, false
		}
		if expression[index] == '"' {
			if index+1 >= len(expression)-1 || expression[index+1] != '"' {
				return 0, false
			}
			index++
		}
		length++
	}
	return length, true
}

func arrayFormalParameterIndexFromSignature(signature procedureSignature, name string) (int, bool) {
	for index, parameter := range signature.Params.AllIndexed() {
		if strings.EqualFold(strings.TrimSpace(parameter.Name), strings.TrimSpace(name)) {
			return index, true
		}
	}
	return 0, false
}

var (
	arrayStringNonEmptyBlockRe = regexp.MustCompile(`(?i)^\s*if\s+(?:[a-z_]\w*\.)?len\s*\(\s*([a-z_]\w*)\s*\)\s*>\s*0\s+then\s*$`)
	arrayStringLengthAssignRe  = regexp.MustCompile(`(?i)^\s*([a-z_]\w*)\s*=\s*(?:[a-z_]\w*\.)?len\s*\(\s*([a-z_]\w*)\s*\)\s*$`)
	arrayStringEmptyExitRe     = regexp.MustCompile(`(?i)^\s*if\s+([a-z_]\w*)\s*=\s*0\s+then\s+exit\s+(?:sub|function|property)\s*$`)
	arrayStringEmptyLenExitRe  = regexp.MustCompile(`(?i)^\s*if\s+(?:[a-z_]\w*\.)?len\s*\(\s*([a-z_]\w*)\s*\)\s*=\s*0\s+then\s+exit\s+(?:sub|function|property)\s*$`)
)

func arrayExpressionKnownNonEmpty(file parsedFile, proc sourceProcedure, line int, rhs string, variables map[string]arrayVariable) bool {
	open := firstParenOutsideString(strings.TrimSpace(rhs))
	if open < 0 {
		return false
	}
	close := matchingParen(rhs, open)
	if close < 0 || strings.TrimSpace(rhs[close+1:]) != "" {
		return false
	}
	for _, argument := range splitArgs(rhs[open+1 : close]) {
		name := strings.ToLower(cleanIdentifier(strings.TrimSpace(argument)))
		variable, ok := variables[name]
		if !ok || variable.isArray || variable.isVariant || !strings.EqualFold(strings.TrimSpace(variable.typ), "String") {
			continue
		}
		if arrayStringIsKnownNonEmpty(file, proc, line, name) {
			return true
		}
	}
	return false
}

// VBA copies a String into a Byte array. A non-empty source establishes a
// usable allocation; vbNullString establishes a known empty allocation whose
// bounds may be queried but whose elements must not be indexed.
func byteArrayStringAssignment(file parsedFile, proc sourceProcedure, line int, variable arrayVariable, rhs string, variables map[string]arrayVariable) (arrayValue, bool) {
	allocated := func(mayBeEmpty bool) (arrayValue, bool) {
		return arrayValue{kind: arrayAllocated, knownArray: true, mayBeEmpty: mayBeEmpty, origin: arrayOriginLocal}, true
	}
	if !variable.isArray || !strings.EqualFold(strings.TrimSpace(variable.typ), "Byte") {
		return arrayValue{}, false
	}
	rhs = strings.TrimSpace(rhs)
	if strings.EqualFold(rhs, "vbNullString") {
		return allocated(true)
	}
	if strings.EqualFold(arrayCallName(rhs), "strconv") {
		open := firstParenOutsideString(rhs)
		close := matchingParen(rhs, open)
		if open < 0 || close < 0 || strings.TrimSpace(rhs[close+1:]) != "" {
			return arrayValue{}, false
		}
		arguments := splitArgs(rhs[open+1 : close])
		if len(arguments) == 0 {
			return arrayValue{}, false
		}
		if strings.EqualFold(strings.TrimSpace(arguments[0]), "vbNullString") {
			return allocated(true)
		}
		if arrayStringExpressionKnownNonEmpty(file, proc, line, arguments[0], variables) {
			return allocated(false)
		}
		return arrayValue{}, false
	}
	if strings.HasPrefix(rhs, `"`) {
		if len(rhs) <= 1 || strings.HasPrefix(rhs, `""`) {
			return arrayValue{}, false
		}
		return allocated(false)
	}
	if !arrayEraseNameRe.MatchString(rhs) {
		return arrayValue{}, false
	}
	source, ok := variables[strings.ToLower(rhs)]
	if !ok || source.isArray || source.isVariant || !strings.EqualFold(strings.TrimSpace(source.typ), "String") {
		return arrayValue{}, false
	}
	if !arrayStringIsKnownNonEmpty(file, proc, line, rhs) {
		return arrayValue{}, false
	}
	return allocated(false)
}

func arrayStringExpressionKnownNonEmpty(file parsedFile, proc sourceProcedure, line int, expression string, variables map[string]arrayVariable) bool {
	expression = strings.TrimSpace(expression)
	if expression == "" || strings.EqualFold(expression, "vbNullString") {
		return false
	}
	if arrayStringExpressionHasNonEmptyLiteral(expression) {
		return true
	}
	name := cleanIdentifier(expression)
	if name != expression {
		return false
	}
	variable, ok := variables[strings.ToLower(name)]
	if !ok || variable.isArray || variable.isVariant || !strings.EqualFold(strings.TrimSpace(variable.typ), "String") {
		return false
	}
	if arrayStringIsKnownNonEmpty(file, proc, line, name) {
		return true
	}
	return arrayStringVariableHasNonEmptyAssignment(file, proc, line, name)
}

func arrayStringExpressionHasNonEmptyLiteral(expression string) bool {
	expression = strings.TrimSpace(expression)
	for strings.HasPrefix(expression, "(") {
		close := matchingParen(expression, 0)
		if close != len(expression)-1 {
			break
		}
		expression = strings.TrimSpace(expression[1:close])
	}
	for _, operand := range splitStringConcatenation(expression) {
		if arrayStringLiteralHasValue(operand) || arrayStringNonEmptyConstant(operand) {
			return true
		}
	}
	return false
}

func arrayStringNonEmptyConstant(operand string) bool {
	switch strings.ToLower(strings.TrimSpace(operand)) {
	case "vbnullchar", "vbcrlf", "vbcr", "vblf", "vbtab", "vbverticaltab", "vbformfeed", "vbnewline":
		return true
	default:
		return false
	}
}

func arrayStringVariableHasNonEmptyAssignment(file parsedFile, proc sourceProcedure, line int, source string) bool {
	start := max(0, proc.StartLine-1)
	end := min(len(file.Lines), line-1)
	depth := 0
	knownNonEmpty := false
	for index := start; index < end; index++ {
		for _, statement := range splitRangeValueSourceStatements(strings.TrimSpace(normalizedCodeLine(file.Lines[index]))) {
			text := strings.TrimSpace(statement)
			if delta := arrayStringBlockBoundary(text); delta < 0 {
				if depth > 0 {
					depth--
				}
				continue
			}
			lhs, rhs, indexed, ok := arrayAssignment(text)
			if ok && !indexed && strings.EqualFold(lhs, source) {
				if nonEmpty := arrayStringExpressionHasNonEmptyLiteral(rhs); depth == 0 {
					knownNonEmpty = nonEmpty
				} else if !nonEmpty {
					// A conditional non-empty assignment preserves an already
					// proven value: the branch either keeps the old value or
					// replaces it with another non-empty value. Any other
					// conditional assignment invalidates the proof.
					knownNonEmpty = false
				}
			}
			if delta := arrayStringBlockBoundary(text); delta > 0 {
				depth++
			}
		}
	}
	return knownNonEmpty
}

func splitStringConcatenation(expression string) []string {
	var parts []string
	start := 0
	inString := false
	depth := 0
	for index := 0; index < len(expression); index++ {
		switch expression[index] {
		case '"':
			if inString && index+1 < len(expression) && expression[index+1] == '"' {
				index++
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
		case '&':
			if !inString && depth == 0 {
				parts = append(parts, strings.TrimSpace(expression[start:index]))
				start = index + 1
			}
		}
	}
	return append(parts, strings.TrimSpace(expression[start:]))
}

func arrayStringLiteralHasValue(operand string) bool {
	operand = strings.TrimSpace(operand)
	if len(operand) < 2 || operand[0] != '"' || operand[len(operand)-1] != '"' {
		return false
	}
	length := 0
	for index := 1; index < len(operand)-1; index++ {
		if operand[index] == '"' {
			if index+1 < len(operand)-1 && operand[index+1] == '"' {
				length++
				index++
				continue
			}
			return false
		}
		length++
	}
	return length > 0
}

// arrayStringBlockBoundary tracks only constructs that can make a String
// assignment conditional. A With block changes name resolution, not whether
// its body executes, so it must not hide an unconditional non-empty value.
func arrayStringBlockBoundary(text string) int {
	lower := strings.ToLower(strings.TrimSpace(text))
	switch {
	case lower == "end if", lower == "end select", lower == "loop", lower == "wend", strings.HasPrefix(lower, "next"):
		return -1
	case strings.HasPrefix(lower, "if ") && strings.HasSuffix(lower, " then"), strings.HasPrefix(lower, "for "), strings.HasPrefix(lower, "do "), lower == "do", strings.HasPrefix(lower, "select case "), strings.HasPrefix(lower, "while "):
		return 1
	default:
		return 0
	}
}

func arrayStringIsKnownNonEmpty(file parsedFile, proc sourceProcedure, line int, source string) bool {
	start := max(0, proc.StartLine-1)
	end := min(len(file.Lines), line-1)
	lengthVariables := map[string]bool{}
	for index := start; index < end; index++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[index]))
		if match := arrayStringEmptyLenExitRe.FindStringSubmatch(text); len(match) == 2 && strings.EqualFold(match[1], source) {
			return true
		}
		if match := arrayStringLengthAssignRe.FindStringSubmatch(text); len(match) == 3 && strings.EqualFold(match[2], source) {
			lengthVariables[strings.ToLower(match[1])] = true
		}
		if match := arrayStringEmptyExitRe.FindStringSubmatch(text); len(match) == 2 && lengthVariables[strings.ToLower(match[1])] {
			return true
		}
	}

	depth := 0
	guardDepth := 0
	for index := start; index < end; index++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[index]))
		lower := strings.ToLower(text)
		if lower == "end if" {
			if guardDepth == depth {
				guardDepth = 0
			}
			if depth > 0 {
				depth--
			}
			continue
		}
		if match := arrayStringNonEmptyBlockRe.FindStringSubmatch(text); len(match) == 2 && strings.EqualFold(match[1], source) {
			depth++
			guardDepth = depth
			continue
		}
		if strings.HasPrefix(lower, "if ") && strings.HasSuffix(lower, " then") {
			depth++
		}
	}
	return guardDepth > 0
}

func arrayCallName(text string) string {
	trimmed := strings.TrimSpace(text)
	if open := firstParenOutsideString(trimmed); open >= 0 {
		if close := matchingParen(trimmed, open); close >= 0 && strings.TrimSpace(trimmed[close+1:]) == "" {
			name := strings.TrimSpace(trimmed[:open])
			if dot := strings.LastIndex(name, "."); dot >= 0 {
				name = name[dot+1:]
			}
			return strings.ToLower(cleanIdentifier(name))
		}
	}
	name := strings.TrimSpace(lastName(trimmed))
	if open := strings.Index(name, "("); open >= 0 {
		name = name[:open]
	}
	return strings.ToLower(cleanIdentifier(name))
}

func firstParenOutsideString(text string) int {
	inString := false
	for i := 0; i < len(text); i++ {
		switch text[i] {
		case '"':
			if inString && i+1 < len(text) && text[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
		case '(':
			if !inString {
				return i
			}
		}
	}
	return -1
}

// inferArrayAllocationGuards recognizes a deliberately small helper
// contract: a scalar Function with one array or Variant parameter returns a
// positive UBound-based length on its normal path and returns zero from an
// On Error GoTo recovery label. That contract is enough to prove the positive
// branch of a direct call, while arbitrary helper functions remain unknown.
func inferArrayAllocationGuards(files []parsedFile) map[string]bool {
	candidates := map[string][]string{}
	procedureNames := map[string]int{}
	recognizedNames := map[string]int{}
	for _, file := range files {
		procedures := file.procedureView()
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			proc := procedures.valueAt(procedureIndex)
			name := strings.ToLower(proc.Name)
			if name != "" {
				procedureNames[name]++
			}
			parameter, ok := arrayAllocationGuardParameter(proc)
			if !ok {
				parameter, ok = arraySafeArrayPointerLengthGuardParameter(file, proc)
			}
			if !ok {
				parameter, ok = arrayDimensionCountGuardParameter(proc)
			}
			if !ok {
				continue
			}
			candidates[name] = append(candidates[name], parameter)
			recognizedNames[name]++
		}
	}
	guards := map[string]bool{}
	for name := range candidates {
		// Private procedures are module-scoped, so the same helper name can
		// legitimately occur in multiple modules. Keep a bare-name guard only
		// when every procedure with that name has the same narrow guard shape;
		// an unrelated duplicate remains conservative.
		if name == "" || recognizedNames[name] != procedureNames[name] {
			continue
		}
		guards[name] = true
	}
	return guards
}

func inferArraySafeArrayLengthGuards(files []parsedFile) map[string]bool {
	procedureNames := map[string]int{}
	recognizedNames := map[string]int{}
	for _, file := range files {
		procedures := file.procedureView()
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			proc := procedures.valueAt(procedureIndex)
			name := strings.ToLower(proc.Name)
			if name == "" {
				continue
			}
			procedureNames[name]++
			if _, ok := arraySafeArrayPointerLengthGuardParameter(file, proc); ok {
				recognizedNames[name]++
			}
		}
	}
	guards := map[string]bool{}
	for name, count := range recognizedNames {
		if count > 0 && count == procedureNames[name] {
			guards[name] = true
		}
	}
	return guards
}

// inferArraySafeBoundGuards recognizes helpers that return UBound(array) on
// their normal path and a negative sentinel after catching an unallocated
// array. A nonnegative result therefore proves both that UBound succeeded and
// that the array has an index at zero or above.
func inferArraySafeBoundGuards(files []parsedFile) map[string]bool {
	candidates := map[string][]string{}
	procedureNames := map[string]int{}
	for _, file := range files {
		procedures := file.procedureView()
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			proc := procedures.valueAt(procedureIndex)
			name := strings.ToLower(proc.Name)
			if name != "" {
				procedureNames[name]++
			}
			parameter, ok := arraySafeBoundGuardParameter(proc)
			if !ok {
				continue
			}
			candidates[name] = append(candidates[name], parameter)
		}
	}
	guards := map[string]bool{}
	for name, parameters := range candidates {
		if name == "" || procedureNames[name] != 1 || len(parameters) != 1 {
			continue
		}
		guards[name] = true
	}
	return guards
}

func arraySafeBoundGuardParameter(proc sourceProcedure) (string, bool) {
	if proc.ProcedureKind != procedureir.ProcedureFunction && proc.ProcedureKind != procedureir.ProcedurePropertyGet {
		return "", false
	}
	if !arrayKnownScalarType(proc.ReturnType) || isObjectType(proc.ReturnType) || proc.Params.Len() != 1 {
		return "", false
	}
	parameter := proc.Params.valueAt(0)
	variantParameter := strings.EqualFold(cleanIdentifier(strings.TrimSpace(parameter.Type)), "variant")
	if parameter.Name == "" || (!parameterIsArray(parameter) && !variantParameter) || proc.Name == "" {
		return "", false
	}

	errorLabel := ""
	recovery := false
	hasRecovery := false
	foundRecoveryLabel := false
	normalReturns := 0
	recoveryNegativeReturns := 0
	invalidReturn := false
	for statement := range proc.Statements.All() {
		text := strings.TrimSpace(normalizedCodeLine(statement.Text))
		if match := arrayOnErrorGotoRe.FindStringSubmatch(text); len(match) == 2 {
			errorLabel = strings.ToLower(match[1])
			hasRecovery = true
		}
		if match := arrayLabelRe.FindStringSubmatch(text); len(match) == 2 {
			recovery = errorLabel != "" && strings.EqualFold(match[1], errorLabel)
			if recovery {
				foundRecoveryLabel = true
			}
		}
		lhs, rhs, indexed, assigned := arrayAssignment(text)
		if !assigned || indexed || !strings.EqualFold(lhs, proc.Name) {
			continue
		}
		switch {
		case arrayUpperBoundExpressionMatches(rhs, parameter.Name):
			normalReturns++
		case recovery && strings.TrimSpace(rhs) == "-1":
			recoveryNegativeReturns++
		default:
			invalidReturn = true
		}
	}
	if !hasRecovery || !foundRecoveryLabel || normalReturns != 1 || recoveryNegativeReturns != 1 || invalidReturn {
		return "", false
	}
	return strings.ToLower(parameter.Name), true
}

func arrayUpperBoundExpressionMatches(rhs, parameter string) bool {
	compact := strings.Join(strings.Fields(strings.ToLower(rhs)), "")
	return compact == "ubound("+strings.ToLower(parameter)+")"
}

func arrayAllocationGuardParameter(proc sourceProcedure) (string, bool) {
	if proc.ProcedureKind != procedureir.ProcedureFunction && proc.ProcedureKind != procedureir.ProcedurePropertyGet {
		return "", false
	}
	if !arrayKnownScalarType(proc.ReturnType) || isObjectType(proc.ReturnType) || proc.Params.Len() != 1 {
		return "", false
	}
	parameter := proc.Params.valueAt(0)
	variantParameter := strings.EqualFold(cleanIdentifier(strings.TrimSpace(parameter.Type)), "variant")
	if parameter.Name == "" || (!parameterIsArray(parameter) && !variantParameter) || proc.Name == "" {
		return "", false
	}

	errorLabel := ""
	recovery := false
	hasRecovery := false
	foundRecoveryLabel := false
	positiveReturns := 0
	recoveryZeroReturns := 0
	invalidReturn := false
	for statement := range proc.Statements.All() {
		text := strings.TrimSpace(normalizedCodeLine(statement.Text))
		if match := arrayOnErrorGotoRe.FindStringSubmatch(text); len(match) == 2 {
			errorLabel = strings.ToLower(match[1])
			hasRecovery = true
		}
		if match := arrayLabelRe.FindStringSubmatch(text); len(match) == 2 {
			recovery = errorLabel != "" && strings.EqualFold(match[1], errorLabel)
			if recovery {
				foundRecoveryLabel = true
			}
		}
		lhs, rhs, indexed, assigned := arrayAssignment(text)
		if !assigned || indexed || !strings.EqualFold(lhs, proc.Name) {
			continue
		}
		switch {
		case arrayLengthExpressionMatches(rhs, parameter.Name):
			positiveReturns++
		case recovery && strings.TrimSpace(rhs) == "0":
			recoveryZeroReturns++
		default:
			invalidReturn = true
		}
	}
	// A typed VBA Function defaults its return value to zero.  A recovery
	// label that falls through to End Function therefore has the same
	// allocation-probe contract as an explicit `FunctionName = 0` assignment,
	// provided the recovery label was actually found and no other return
	// assignment invalidated the shape above.
	if !hasRecovery || !foundRecoveryLabel || positiveReturns != 1 || recoveryZeroReturns > 1 || invalidReturn {
		return "", false
	}
	return strings.ToLower(parameter.Name), true
}

// arraySafeArrayPointerLengthGuardParameter recognizes a scalar helper that
// returns the length of a Byte array after the low-level SAFEARRAY descriptor
// guard. Unlike the ordinary On Error-based allocation probe, this form uses
// the function's default zero return on either early-exit path. Require the
// returned expression to be derived from preceding LBound and UBound
// assignments so an arbitrary pointer check cannot become an allocation
// contract.
func arraySafeArrayPointerLengthGuardParameter(file parsedFile, proc sourceProcedure) (string, bool) {
	if proc.ProcedureKind != procedureir.ProcedureFunction && proc.ProcedureKind != procedureir.ProcedurePropertyGet {
		return "", false
	}
	if !arrayKnownScalarType(proc.ReturnType) || isObjectType(proc.ReturnType) || proc.Params.Len() != 1 {
		return "", false
	}
	parameter := proc.Params.valueAt(0)
	if parameter.Name == "" || !parameterIsArray(parameter) || proc.Name == "" {
		return "", false
	}
	variables := arrayVariables(file, proc, file.moduleDecls())
	parameterName := strings.ToLower(cleanIdentifier(parameter.Name))
	variable, known := variables[parameterName]
	if !known || !variable.isArray || !isByteArrayVariable(variable) {
		return "", false
	}
	guardLine := 0
	for line := proc.StartLine; line <= proc.EndLine && line <= len(file.Lines); line++ {
		if target, ok := arraySafeArrayPointerGuardTarget(file, proc, line, file.Lines[line-1], variables); ok && target == parameterName {
			guardLine = line
			break
		}
	}
	if guardLine == 0 {
		return "", false
	}
	lowerName := ""
	upperName := ""
	returnCount := 0
	invalidReturn := false
	for line := guardLine + 1; line <= proc.EndLine && line <= len(file.Lines); line++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[line-1]))
		if text == "" || strings.HasPrefix(text, "'") || strings.HasPrefix(text, "#") {
			continue
		}
		lhs, rhs, indexed, assigned := arrayAssignment(text)
		if !assigned || indexed {
			continue
		}
		compact := strings.Join(strings.Fields(strings.ToLower(rhs)), "")
		switch {
		case compact == "lbound("+parameterName+")":
			if lowerName != "" {
				return "", false
			}
			lowerName = strings.ToLower(cleanIdentifier(lhs))
		case compact == "ubound("+parameterName+")":
			if upperName != "" {
				return "", false
			}
			upperName = strings.ToLower(cleanIdentifier(lhs))
		case strings.EqualFold(cleanIdentifier(lhs), proc.Name):
			returnCount++
			expected := upperName + "-" + lowerName + "+1"
			if lowerName == "" || upperName == "" || compact != expected {
				invalidReturn = true
			}
		}
	}
	if lowerName == "" || upperName == "" || returnCount != 1 || invalidReturn {
		return "", false
	}
	return parameterName, true
}

// arrayDimensionCountGuardParameter recognizes the helper shape used by
// GetArrayDimsCount: it probes successive LBound dimensions under an error
// handler and returns the last successful dimension number minus one. The
// result is zero when the first probe fails, so a caller branch that proves
// the result is one has also proved that its input is an allocated 1D array.
// Keep this separate from the ordinary length probe so an arbitrary scalar
// helper returning `someValue - 1` cannot become an allocation contract.
func arrayDimensionCountGuardParameter(proc sourceProcedure) (string, bool) {
	if proc.ProcedureKind != procedureir.ProcedureFunction && proc.ProcedureKind != procedureir.ProcedurePropertyGet {
		return "", false
	}
	if !arrayKnownScalarType(proc.ReturnType) || isObjectType(proc.ReturnType) || proc.Params.Len() != 1 {
		return "", false
	}
	parameter := proc.Params.valueAt(0)
	variantParameter := strings.EqualFold(cleanIdentifier(strings.TrimSpace(parameter.Type)), "variant")
	if parameter.Name == "" || (!parameterIsArray(parameter) && !variantParameter) || proc.Name == "" {
		return "", false
	}

	errorLabel := ""
	recovery := false
	foundRecoveryLabel := false
	loopVariable := ""
	hasDimensionProbe := false
	countReturns := 0
	invalidReturn := false
	for statement := range proc.Statements.All() {
		rawText := strings.TrimSpace(statement.Text)
		text := strings.TrimSpace(normalizedCodeLine(statement.Text))
		if match := arrayOnErrorGotoRe.FindStringSubmatch(text); len(match) == 2 {
			errorLabel = strings.ToLower(match[1])
		}
		if match := arrayLabelRe.FindStringSubmatch(text); len(match) == 2 {
			recovery = errorLabel != "" && strings.EqualFold(match[1], errorLabel)
			if recovery {
				foundRecoveryLabel = true
			}
		}
		loopText := rawText
		if newline := strings.IndexAny(loopText, "\r\n"); newline >= 0 {
			loopText = strings.TrimSpace(loopText[:newline])
		}
		if match := arrayDimensionCountLoopRe.FindStringSubmatch(loopText); len(match) == 2 {
			loopVariable = strings.ToLower(cleanIdentifier(match[1]))
		}
		for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(text, -1) {
			if !strings.EqualFold(bound[1], "lbound") || !strings.EqualFold(cleanIdentifier(bound[2]), parameter.Name) || loopVariable == "" || !strings.EqualFold(cleanIdentifier(bound[3]), loopVariable) {
				continue
			}
			hasDimensionProbe = true
		}
		lhs, rhs, indexed, assigned := arrayAssignment(text)
		if !assigned || indexed || !strings.EqualFold(lhs, proc.Name) {
			continue
		}
		if recovery && arrayDimensionCountExpressionMatches(rhs, loopVariable) {
			countReturns++
		} else {
			invalidReturn = true
		}
	}
	if errorLabel == "" || !foundRecoveryLabel || loopVariable == "" || !hasDimensionProbe || countReturns != 1 || invalidReturn {
		return "", false
	}
	return strings.ToLower(parameter.Name), true
}

func arrayDimensionCountExpressionMatches(rhs, dimension string) bool {
	compact := strings.Join(strings.Fields(strings.ToLower(rhs)), "")
	return compact == strings.ToLower(dimension)+"-1"
}

func parameterIsArray(parameter parameterInfo) bool {
	return parameter.ParamArray || strings.Contains(parameter.Type, "()") || parameter.ValueShape == procedureir.ValueShapeFixedArray || parameter.ValueShape == procedureir.ValueShapeDynamicArray
}

func arrayLengthExpressionMatches(rhs, parameter string) bool {
	compact := strings.Join(strings.Fields(strings.ToLower(rhs)), "")
	parameter = strings.ToLower(parameter)
	return compact == "ubound("+parameter+")-lbound("+parameter+")+1" || compact == "ubound("+parameter+")+1"
}

// inferArrayReturnSummaries intentionally summarizes only normal, directly
// observed return assignments. Recognized allocation guards refine the normal
// branch, while a definitely failing ReDim without local error handling does
// not contribute a normal path. A missing assignment, mixed assignment kinds,
// duplicate procedure names, and recursive/external calls remain unknown.
// Dependencies are solved to a fixed point so a private array-returning helper
// may be declared after its caller without turning an unproved VBA function
// contract into an allocation guarantee.
func arrayReturnSummaryDuplicateNames(procedures []sourceProcedure) map[string]bool {
	counts := make(map[string]int, len(procedures))
	for _, procedure := range procedures {
		name := strings.ToLower(strings.TrimSpace(procedure.Name))
		if name != "" {
			counts[name]++
		}
	}
	duplicates := make(map[string]bool)
	for name, count := range counts {
		if count > 1 {
			duplicates[name] = true
		}
	}
	return duplicates
}

func inferDocumentedArrayReturnSummaries(files []parsedFile) map[string]arrayValue {
	returns := map[string]arrayValue{}
	for _, file := range files {
		for procedure := range file.procedureView().All() {
			if procedure.ProcedureKind != procedureir.ProcedureFunction && procedure.ProcedureKind != procedureir.ProcedurePropertyGet || procedure.Name == "" {
				continue
			}
			if !arrayProcedureDocumentsArray(file, procedure) || !arrayProcedureHasReturnAllocation(file, procedure) {
				continue
			}
			returns[arrayProcedureKey(procedure)] = arrayValue{
				kind:       arrayAllocated,
				knownArray: true,
				mayBeEmpty: true,
				origin:     arrayOriginLocal,
			}
		}
	}
	return returns
}

func inferDocumentedNonEmptyArrayReturnSummaries(files []parsedFile) map[string]arrayValue {
	returns := map[string]arrayValue{}
	for _, file := range files {
		for procedure := range file.procedureView().All() {
			if procedure.ProcedureKind != procedureir.ProcedureFunction && procedure.ProcedureKind != procedureir.ProcedurePropertyGet || procedure.Name == "" {
				continue
			}
			if !arrayProcedureDocumentsArray(file, procedure) || !arrayProcedureHasNonEmptyReturnAllocation(file, procedure) {
				continue
			}
			returns[arrayProcedureKey(procedure)] = arrayValue{
				kind:       arrayAllocated,
				knownArray: true,
				origin:     arrayOriginLocal,
			}
		}
	}
	return returns
}

// inferDocumentedArrayReturnLowerBoundSummaries records a weaker contract for
// documented helpers that may return an unallocated array on an empty input,
// but consistently allocate that array from a known lower bound. The caller
// still receives a VBA227 for a direct UBound on the possibly-unallocated
// result; the lower bound is used only after that query has successfully
// reached a For body.
func inferDocumentedArrayReturnLowerBoundSummaries(files []parsedFile) map[string]arrayValue {
	returns := map[string]arrayValue{}
	for _, file := range files {
		for procedure := range file.procedureView().All() {
			if procedure.ProcedureKind != procedureir.ProcedureFunction && procedure.ProcedureKind != procedureir.ProcedurePropertyGet || procedure.Name == "" {
				continue
			}
			if !arrayProcedureDocumentsArray(file, procedure) || !arrayProcedureHasReturnAllocation(file, procedure) || arrayProcedureHasNonEmptyReturnAllocation(file, procedure) {
				continue
			}
			name, ok := arrayProcedureReturnSource(file, procedure)
			if !ok {
				continue
			}
			lower, ok := arrayProcedureReturnLowerBound(file, procedure, name)
			if !ok {
				continue
			}
			shape := []arrayDimension{{lower: lower}}
			returns[arrayProcedureKey(procedure)] = arrayValue{
				// Keep this as an unknown allocation state. The lower bound is a
				// separate proof used only after a successful UBound query; marking
				// it as an unallocated known array would make the deterministic
				// runtime lane report the same query and suppress the lifecycle
				// diagnostic that the caller still needs to see.
				kind:          arrayUnknown,
				knownArray:    false,
				dimensions:    shape,
				preserveShape: shape,
				origin:        arrayOriginLocal,
			}
		}
	}
	return returns
}

func arrayProcedureReturnSource(file parsedFile, proc sourceProcedure) (string, bool) {
	name := ""
	start := max(0, proc.StartLine-1)
	end := min(proc.EndLine, len(file.Lines))
	for _, rawLine := range file.Lines[start:end] {
		lhs, rhs, indexed, assigned := arrayAssignment(strings.TrimSpace(normalizedCodeLine(rawLine)))
		if !assigned || indexed || !strings.EqualFold(lhs, proc.Name) {
			continue
		}
		source := directArrayArgumentName(rhs)
		if source == "" || name != "" && !strings.EqualFold(name, source) {
			return "", false
		}
		name = source
	}
	return name, name != ""
}

func arrayProcedureReturnLowerBound(file parsedFile, proc sourceProcedure, source string) (arrayBound, bool) {
	var lower arrayBound
	hasLower := false
	start := max(0, proc.StartLine-1)
	end := min(proc.EndLine, len(file.Lines))
	for _, rawLine := range file.Lines[start:end] {
		text := strings.TrimSpace(normalizedCodeLine(rawLine))
		if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 {
			for _, clause := range splitArgs(match[2]) {
				redim, direct := parseDirectArrayRedimClause(clause)
				if !direct || !strings.EqualFold(cleanIdentifier(redim.name), source) {
					continue
				}
				dimensions := parseArrayDimensionsWithConstants(redim.dimensions, arrayOptionBase(file), arrayIntegerConstants(file, proc, nil, nil))
				if len(dimensions) != 1 || !dimensions[0].lower.known {
					return arrayBound{}, false
				}
				if hasLower && lower.value != dimensions[0].lower.value {
					return arrayBound{}, false
				}
				lower = dimensions[0].lower
				hasLower = true
			}
		}
		if match := arrayEraseRe.FindStringSubmatch(text); len(match) == 2 {
			for _, target := range splitArgs(match[1]) {
				if strings.EqualFold(cleanIdentifier(target), source) {
					return arrayBound{}, false
				}
			}
		}
		if lhs, _, indexed, assigned := arrayAssignment(text); assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), source) {
			return arrayBound{}, false
		}
	}
	return lower, hasLower
}

func arrayProcedureDocumentsArray(file parsedFile, proc sourceProcedure) bool {
	start := max(0, proc.StartLine-1-5)
	end := min(max(0, proc.StartLine-1), len(file.Lines))
	for _, rawLine := range file.Lines[start:end] {
		comment := strings.TrimSpace(rawLine)
		if !strings.HasPrefix(comment, "'") {
			continue
		}
		comment = strings.TrimSpace(comment[1:])
		if arrayReturnArrayDocRe.MatchString(comment) {
			return true
		}
	}
	return false
}

func arrayProcedureHasReturnAllocation(file parsedFile, proc sourceProcedure) bool {
	hasAllocation := false
	hasReturn := false
	start := max(0, proc.StartLine-1)
	end := min(proc.EndLine, len(file.Lines))
	for _, rawLine := range file.Lines[start:end] {
		text := strings.TrimSpace(normalizedCodeLine(rawLine))
		lower := strings.ToLower(text)
		if strings.HasPrefix(lower, "redim ") {
			hasAllocation = true
		}
		lhs, _, indexed, assigned := arrayAssignment(text)
		if assigned && !indexed && strings.EqualFold(lhs, proc.Name) {
			hasReturn = true
		}
	}
	return hasAllocation && hasReturn
}

// arrayDescriptorArrayReturnSummary recognizes a typed array Function that
// returns a view backed by a persistent SAFEARRAY descriptor. The contract is
// intentionally structural: a Static UDT-like accessor is initialized behind
// its readiness flag, the descriptor data/count/lower-bound fields come from
// scalar parameters, and the direct array member is returned on every normal
// path. Unknown call arguments retain a possible-empty allocation; known
// literal arguments can recover the returned one-dimensional shape later.
func arrayDescriptorArrayReturnSummary(file parsedFile, proc sourceProcedure, ctx analysisContext) (arrayValue, bool) {
	if proc.Graph == nil || proc.ProcedureKind != procedureir.ProcedureFunction && proc.ProcedureKind != procedureir.ProcedurePropertyGet {
		return arrayValue{}, false
	}
	if proc.ReturnValueShape != procedureir.ValueShapeDynamicArray && !strings.Contains(strings.ReplaceAll(proc.ReturnType, " ", ""), "()") {
		return arrayValue{}, false
	}
	root, returnLine, ok := arrayDescriptorReturnSource(file, proc)
	if !ok {
		return arrayValue{}, false
	}
	variables := arrayVariables(file, proc, file.moduleDecls())
	ready, known := variables[strings.ToLower(root)]
	if !known || !ready.static || ready.isArray || ready.isVariant || ready.isObject || ready.knownScalar {
		return arrayValue{}, false
	}
	if !arrayDescriptorReadyInitialized(file, proc, root, ctx) {
		return arrayValue{}, false
	}
	source, start, length, lower, setupLines, ok := arrayDescriptorReturnSetup(file, proc, root)
	if !ok || !arrayDescriptorParameter(proc, source, "String") || !arrayDescriptorParameter(proc, start, "") || !arrayDescriptorParameter(proc, length, "") || !arrayDescriptorParameter(proc, lower, "") {
		return arrayValue{}, false
	}
	for _, line := range setupLines {
		if line >= returnLine || !arrayDescriptorLineDominatesNormalExit(proc, line) {
			return arrayValue{}, false
		}
	}
	if !proc.Graph.IsDefinitelyAssigned(proc.Graph.NormalExit, vbacfg.Variable{Scope: procedureir.ScopeLocal, Name: proc.Name}, vbacfg.EdgeFilter{NormalOnly: true}) {
		return arrayValue{}, false
	}
	return arrayValue{
		kind:                            arrayAllocated,
		knownArray:                      true,
		mayBeEmpty:                      true,
		origin:                          arrayOriginLocal,
		returnDescriptorSourceParameter: strings.ToLower(source),
		returnDescriptorStartParameter:  strings.ToLower(start),
		returnDescriptorLengthParameter: strings.ToLower(length),
		returnDescriptorLowerParameter:  strings.ToLower(lower),
	}, true
}

func arrayDescriptorReturnSource(file parsedFile, proc sourceProcedure) (string, int, bool) {
	root := ""
	returnLine := 0
	for line := max(0, proc.StartLine-1); line < min(proc.EndLine, len(file.Lines)); line++ {
		lhs, rhs, indexed, assigned := arrayAssignment(strings.TrimSpace(normalizedCodeLine(file.Lines[line])))
		if !assigned || indexed || !strings.EqualFold(lhs, proc.Name) {
			continue
		}
		if returnLine != 0 {
			return "", 0, false
		}
		receiver, member, ok := arrayQualifiedMemberParts(rhs)
		if !ok || receiver == "" {
			return "", 0, false
		}
		parts := append(strings.Split(receiver, "."), member)
		if len(parts) != 3 || !strings.EqualFold(parts[1], "ac") || !strings.HasPrefix(strings.ToLower(parts[2]), "d") {
			return "", 0, false
		}
		root = parts[0]
		returnLine = line + 1
	}
	return root, returnLine, root != "" && returnLine != 0
}

func arrayDescriptorReadyInitialized(file parsedFile, proc sourceProcedure, root string, ctx analysisContext) bool {
	pattern := regexp.MustCompile(`(?i)(?:^|:)\s*if\s+not\s+` + regexp.QuoteMeta(root) + `\s*\.\s*isset\s+then\s+initmemoryaccessor\s+` + regexp.QuoteMeta(root) + `\s*$`)
	guardLine := -1
	for line := max(0, proc.StartLine-1); line < min(proc.EndLine, len(file.Lines)); line++ {
		if pattern.MatchString(strings.TrimSpace(normalizedCodeLine(file.Lines[line]))) {
			if guardLine >= 0 {
				return false
			}
			guardLine = line
		}
	}
	if guardLine < 0 {
		return false
	}
	initializerFound := false
	for call := range proc.Calls.All() {
		if call.Range.StartLine-1 != guardLine {
			continue
		}
		if initializerFound {
			return false
		}
		helper, parameter, resolved := arrayStaticReadyInitializer(file, proc, call, root, ctx)
		if !resolved || helper.StartByte == proc.StartByte || !arrayStaticHelperSetsReadyFlag(file, helper, parameter) {
			return false
		}
		initializerFound = true
	}
	return initializerFound
}

func arrayDescriptorReturnSetup(file parsedFile, proc sourceProcedure, root string) (string, string, string, string, []int, bool) {
	prefix := regexp.QuoteMeta(root)
	pvRe := regexp.MustCompile(`(?i)^\s*` + prefix + `\s*\.\s*sa\s*\.\s*pvdata\s*=\s*strptr\s*\(\s*([A-Za-z_]\w*)\s*\)\s*\+\s*\(\s*([A-Za-z_]\w*)\s*-\s*1\s*\)\s*\*\s*[A-Za-z_]\w*\s*$`)
	cbRe := regexp.MustCompile(`(?i)^\s*` + prefix + `\s*\.\s*sa\s*\.\s*cbelements\s*=\s*\S.*$`)
	lowerRe := regexp.MustCompile(`(?i)^\s*` + prefix + `\s*\.\s*sa\s*\.\s*rgsabound0\s*\.\s*llbound\s*=\s*([A-Za-z_]\w*)\s*$`)
	countRe := regexp.MustCompile(`(?i)^\s*` + prefix + `\s*\.\s*sa\s*\.\s*rgsabound0\s*\.\s*celements\s*=\s*([A-Za-z_]\w*)\s*$`)
	source, start, length, lower := "", "", "", ""
	pvCount, cbCount, lowerCount, countCount := 0, 0, 0, 0
	lines := make([]int, 0, 4)
	for line := max(0, proc.StartLine-1); line < min(proc.EndLine, len(file.Lines)); line++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[line]))
		if match := pvRe.FindStringSubmatch(text); len(match) == 3 {
			pvCount++
			source, start = match[1], match[2]
			lines = append(lines, line+1)
		}
		if cbRe.MatchString(text) {
			cbCount++
			lines = append(lines, line+1)
		}
		if match := lowerRe.FindStringSubmatch(text); len(match) == 2 {
			lowerCount++
			lower = match[1]
			lines = append(lines, line+1)
		}
		if match := countRe.FindStringSubmatch(text); len(match) == 2 {
			countCount++
			length = match[1]
			lines = append(lines, line+1)
		}
	}
	if pvCount != 1 || cbCount != 1 || lowerCount != 1 || countCount != 1 {
		return "", "", "", "", nil, false
	}
	return source, start, length, lower, lines, true
}

func arrayDescriptorParameter(proc sourceProcedure, name, typeName string) bool {
	for parameter := range proc.Params.All() {
		if !strings.EqualFold(strings.TrimSpace(parameter.Name), strings.TrimSpace(name)) {
			continue
		}
		if parameter.IsArray || parameter.ValueShape == procedureir.ValueShapeFixedArray || parameter.ValueShape == procedureir.ValueShapeDynamicArray {
			return false
		}
		return typeName == "" || strings.EqualFold(strings.TrimSpace(parameter.Type), typeName)
	}
	return false
}

func arrayDescriptorLineDominatesNormalExit(proc sourceProcedure, line int) bool {
	dominators := arrayProcedureNormalExitDominators(proc)
	for statement := range proc.Statements.All() {
		if statement.Range.StartLine == line && arrayProcedureBlockDominatesNormalExit(proc, statement.ID, dominators) {
			return true
		}
	}
	return false
}

// arrayConditionalReturnSummary recognizes a small, path-sensitive family of
// array factories. The returned array is intentionally kept conditional: the
// caller must prove either that the scalar length is positive or that the
// input Byte array is non-empty before this fact becomes an allocation proof.
// Requiring one direct return source, one guarded ReDim, and a definitely
// assigned function result keeps this summary fail-open for general helper
// functions and error paths.
func arrayConditionalReturnSummary(file parsedFile, proc sourceProcedure) (arrayValue, bool) {
	if proc.Graph == nil || proc.ProcedureKind != procedureir.ProcedureFunction && proc.ProcedureKind != procedureir.ProcedurePropertyGet {
		return arrayValue{}, false
	}
	source, ok := arrayProcedureReturnSource(file, proc)
	if !ok {
		return arrayValue{}, false
	}
	source = strings.ToLower(cleanIdentifier(source))
	variables := arrayVariables(file, proc, file.moduleDecls())
	returnedVariable, ok := variables[source]
	if !ok || !returnedVariable.isArray || returnedVariable.fixed || returnedVariable.parameter {
		return arrayValue{}, false
	}
	if !proc.Graph.IsDefinitelyAssigned(proc.Graph.NormalExit, vbacfg.Variable{Scope: procedureir.ScopeLocal, Name: proc.Name}, vbacfg.EdgeFilter{NormalOnly: true}) {
		return arrayValue{}, false
	}

	returnLines := make([]int, 0, 1)
	foundConditionalOutput := false
	redimLine := 0
	returnNonEmptyArrayParameter := ""
	returnPositiveScalarParameter := ""
	for statement := range proc.Statements.All() {
		text := strings.TrimSpace(normalizedCodeLine(statement.Text))
		if lhs, _, indexed, assigned := arrayAssignment(text); assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), proc.Name) {
			returnLines = append(returnLines, statement.Range.StartLine)
			continue
		}
		if statement.Kind != procedureir.StatementReDim {
			if lhs, _, indexed, assigned := arrayAssignment(text); assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), source) {
				condition, branch, guarded := arrayConditionalReturnGuard(proc, statement)
				input, proven := arrayConditionalReturnStrPtrInput(condition, branch, variables)
				if !guarded || !proven || returnPositiveScalarParameter != "" || returnNonEmptyArrayParameter != "" && !strings.EqualFold(returnNonEmptyArrayParameter, input) {
					return arrayValue{}, false
				}
				returnNonEmptyArrayParameter = input
				foundConditionalOutput = true
			}
			if match := arrayEraseRe.FindStringSubmatch(text); len(match) == 2 {
				for _, target := range splitArgs(match[1]) {
					if strings.EqualFold(cleanIdentifier(target), source) {
						return arrayValue{}, false
					}
				}
			}
			continue
		}
		match := arrayRedimRe.FindStringSubmatch(text)
		if len(match) == 0 {
			continue
		}
		for _, clause := range splitArgs(match[2]) {
			redim, direct := parseDirectArrayRedimClause(clause)
			if !direct || !strings.EqualFold(cleanIdentifier(redim.name), source) {
				continue
			}
			if strings.TrimSpace(match[1]) != "" {
				return arrayValue{}, false
			}
			condition, branch, guarded := arrayConditionalReturnGuard(proc, statement)
			if !guarded {
				return arrayValue{}, false
			}
			if input, proven := arrayConditionalReturnStrPtrInput(condition, branch, variables); proven {
				if returnPositiveScalarParameter != "" || returnNonEmptyArrayParameter != "" && !strings.EqualFold(returnNonEmptyArrayParameter, input) {
					return arrayValue{}, false
				}
				returnNonEmptyArrayParameter = input
			} else if scalar, proven := arrayConditionalReturnPositiveScalar(condition, branch, redim.dimensions, variables); proven {
				if returnNonEmptyArrayParameter != "" || returnPositiveScalarParameter != "" && !strings.EqualFold(returnPositiveScalarParameter, scalar) {
					return arrayValue{}, false
				}
				returnPositiveScalarParameter = scalar
			} else {
				return arrayValue{}, false
			}
			foundConditionalOutput = true
			if redimLine == 0 || statement.Range.StartLine < redimLine {
				redimLine = statement.Range.StartLine
			}
		}
	}
	if len(returnLines) != 1 || !foundConditionalOutput || redimLine > returnLines[0] {
		return arrayValue{}, false
	}
	return arrayValue{
		kind:                          arrayUnknown,
		knownArray:                    true,
		mayBeEmpty:                    true,
		origin:                        arrayOriginLocal,
		returnNonEmptyArrayParameter:  returnNonEmptyArrayParameter,
		returnPositiveScalarParameter: returnPositiveScalarParameter,
	}, true
}

func arrayConditionalReturnGuard(proc sourceProcedure, statement procedureir.Statement) (string, vbacfg.EdgeKind, bool) {
	current := statement
	visited := map[int]bool{}
	for current.ParentID != 0 && !visited[current.ParentID] {
		visited[current.ParentID] = true
		parent := procedureStatementByID(proc, current.ParentID)
		if parent.ID == 0 {
			return "", "", false
		}
		switch parent.Kind {
		case procedureir.StatementIf, procedureir.StatementElseIf:
			return arrayConditionalReturnCondition(parent), vbacfg.EdgeBranchTrue, true
		case procedureir.StatementElse:
			branch := procedureStatementByID(proc, parent.ParentID)
			if branch.Kind != procedureir.StatementIf && branch.Kind != procedureir.StatementElseIf {
				return "", "", false
			}
			return arrayConditionalReturnCondition(branch), vbacfg.EdgeBranchFalse, true
		}
		current = parent
	}
	return "", "", false
}

func arrayConditionalReturnCondition(statement procedureir.Statement) string {
	condition := statement.Text
	if statement.Condition != nil && strings.TrimSpace(statement.Condition.Text) != "" {
		condition = statement.Condition.Text
	}
	if parsed, _, ok := arrayIfThenParts(condition); ok {
		condition = parsed
	}
	condition = strings.TrimSpace(condition)
	lower := strings.ToLower(condition)
	if strings.HasPrefix(lower, "if ") {
		condition = strings.TrimSpace(condition[3:])
	} else if strings.HasPrefix(lower, "elseif ") {
		condition = strings.TrimSpace(condition[len("elseif "):])
	}
	if then := strings.LastIndex(strings.ToLower(condition), " then"); then >= 0 && strings.TrimSpace(condition[then+5:]) == "" {
		condition = strings.TrimSpace(condition[:then])
	}
	for len(condition) >= 2 && condition[0] == '(' && condition[len(condition)-1] == ')' {
		condition = strings.TrimSpace(condition[1 : len(condition)-1])
	}
	return condition
}

func arrayConditionalReturnStrPtrInput(condition string, branch vbacfg.EdgeKind, variables map[string]arrayVariable) (string, bool) {
	match := arrayStrPtrGuardRe.FindStringSubmatch(strings.TrimSpace(condition))
	if len(match) != 3 {
		return "", false
	}
	name := strings.ToLower(cleanIdentifier(match[1]))
	variable, known := variables[name]
	if !known || !variable.parameter || !parameterIsByRefArrayForVariable(variable) || !isByteArrayVariable(variable) {
		return "", false
	}
	required := vbacfg.EdgeBranchFalse
	if match[2] == "<>" {
		required = vbacfg.EdgeBranchTrue
	}
	return name, branch == required
}

func parameterIsByRefArrayForVariable(variable arrayVariable) bool {
	return variable.parameter && variable.isArray && !variable.paramArray
}

func arrayConditionalReturnPositiveScalar(condition string, branch vbacfg.EdgeKind, dimensions string, variables map[string]arrayVariable) (string, bool) {
	lhs, operator, literal, ok := arrayCountComparison(condition)
	if !ok {
		return "", false
	}
	name := strings.ToLower(cleanIdentifier(lhs))
	variable, known := variables[name]
	if !known || !variable.parameter || variable.isArray || variable.isVariant || variable.isObject {
		return "", false
	}
	positiveBranch, ok := positiveArrayCountBranch(operator, literal)
	if !ok || branch != positiveBranch || !arrayRedimUsesPositiveScalar(dimensions, name) {
		return "", false
	}
	return name, true
}

func arrayRedimUsesPositiveScalar(dimensions, parameter string) bool {
	wanted := strings.ToLower(cleanIdentifier(parameter)) + "-1"
	for _, dimension := range splitArgs(dimensions) {
		compact := strings.ToLower(canonicalArrayBoundExpression(dimension))
		if strings.HasPrefix(compact, "0to") && strings.TrimPrefix(compact, "0to") == wanted {
			return true
		}
	}
	return false
}

// arrayProcedureHasNonEmptyReturnAllocation recognizes the stronger CFG-backed
// contract needed when a documented array return is used by a bare caller.
// A plain ReDim with fully known, non-empty bounds is different from a
// ReDim Preserve inside a loop: the latter can leave the returned array
// unallocated when the loop has no iterations. Every normal return path must
// therefore assign the documented return value from a known non-empty array.
func arrayProcedureHasNonEmptyReturnAllocation(file parsedFile, proc sourceProcedure) bool {
	if proc.Graph == nil {
		return false
	}
	variables := arrayVariables(file, proc, file.moduleDecls())
	constants := arrayIntegerConstants(file, proc, nil, nil)
	ctx := analysisContext{arrayAllowVariantRedim: true}
	type returnCandidate struct {
		value arrayValue
		ok    bool
	}
	returnCandidates := map[int]returnCandidate{}
	base := arrayOptionBase(file)
	procedureHasErrorHandling := arrayProcedureHasErrorHandling(proc)
	// An Err.Raise branch is not a normal return path. Keep exceptional edges
	// available for procedures with active error handling, but remove the
	// synthetic normal continuation so an error-only branch cannot make the
	// return assignment look non-definite.
	graph := proc.Graph.WithoutNormalErrRaiseContinuationView()
	walkArrayCFGWithStopStats(&graph, file.Lines, arrayInitialState(variables), func(text string, line int, in arrayFlowState) arrayFlowState {
		if lhs, rhs, indexed, assigned := arrayAssignment(text); assigned && !indexed && strings.EqualFold(lhs, proc.Name) {
			value, known := arrayExpressionState(rhs, in, ctx)
			returnCandidates[line] = returnCandidate{
				value: value,
				ok:    known && value.kind == arrayAllocated && value.knownArray && value.origin != arrayOriginRangeValue,
			}
		}
		out, _ := (Analyzer{}).arrayTransfer(file, proc, ctx, variables, in, text, line, constants, nil)
		return out
	}, nil, func(text string, _ int) bool {
		return !procedureHasErrorHandling && arraySummaryStatementAlwaysFails(text, base, constants)
	}, nil)
	if len(returnCandidates) == 0 || !proc.Graph.IsDefinitelyAssigned(proc.Graph.NormalExit, vbacfg.Variable{Scope: procedureir.ScopeLocal, Name: proc.Name}, vbacfg.EdgeFilter{NormalOnly: true, WithoutNormalErrRaiseContinuation: true}) {
		return false
	}
	for _, candidate := range returnCandidates {
		if !candidate.ok || candidate.value.mayBeEmpty {
			return false
		}
	}
	return true
}

type arrayReturnSummarySet struct {
	bare      map[string]arrayValue
	qualified map[string]arrayValue
}

func inferArrayReturnSummaries(files []parsedFile, arrayAllocationGuards map[string]bool, participantCtx analysisContext) map[string]arrayValue {
	return inferArrayReturnSummarySet(files, arrayAllocationGuards, participantCtx).bare
}

func inferArrayReturnSummarySet(files []parsedFile, arrayAllocationGuards map[string]bool, participantCtx analysisContext) arrayReturnSummarySet {
	type candidate struct {
		value arrayValue
		ok    bool
	}
	type returnProcedure struct {
		file      parsedFile
		proc      sourceProcedure
		variables map[string]arrayVariable
		constants map[string]int
	}
	procedures := make([]returnProcedure, 0)
	allReturnProcedures := make([]sourceProcedure, 0)
	for _, file := range files {
		fileProcedures := file.procedureView()
		moduleDecls := file.moduleDecls()
		for procedureIndex := 0; procedureIndex < fileProcedures.Len(); procedureIndex++ {
			proc := fileProcedures.valueAt(procedureIndex)
			if proc.ProcedureKind != procedureir.ProcedureFunction && proc.ProcedureKind != procedureir.ProcedurePropertyGet {
				continue
			}
			if proc.Name == "" {
				continue
			}
			// Duplicate bare-name summaries are ambiguous even when one of the
			// same-named procedures is outside the array participant closure.
			// Collect names before filtering so narrowing the fixed-point scope
			// cannot turn an otherwise ambiguous call into a definite summary.
			allReturnProcedures = append(allReturnProcedures, proc)
			if !arrayProcedureIsParticipant(participantCtx, proc) {
				continue
			}
			procedures = append(procedures, returnProcedure{
				file:      file,
				proc:      proc,
				variables: arrayVariables(file, proc, moduleDecls),
				constants: arrayIntegerConstants(file, proc, nil, nil),
			})
		}
	}
	sort.SliceStable(procedures, func(i, j int) bool {
		return arrayProcedureLess(procedures[i].proc, procedures[j].proc)
	})
	ambiguousReturnNames := arrayReturnSummaryDuplicateNames(allReturnProcedures)

	evaluate := func(procedure returnProcedure, summaries map[string]arrayValue) (candidate, bool) {
		proc := procedure.proc
		if proc.Graph == nil {
			// Without a CFG the scan cannot distinguish conditional from
			// unconditional return assignments, so leave the summary unknown.
			return candidate{}, false
		}
		arrayReturns := summaries
		// Never use a previous summary for the procedure currently being
		// inspected. This keeps direct and mutual recursion fail-open while
		// still allowing independent helpers to converge over later rounds.
		if _, self := summaries[strings.ToLower(proc.Name)]; self {
			arrayReturns = make(map[string]arrayValue, len(summaries)-1)
			for name, value := range summaries {
				if !strings.EqualFold(name, proc.Name) {
					arrayReturns[name] = value
				}
			}
		}
		ctx := analysisContext{
			arrayReturns:           arrayReturns,
			arrayAllowVariantRedim: arrayProcedureDocumentsArray(procedure.file, proc) && arrayProcedureHasReturnAllocation(procedure.file, proc),
		}
		returnCandidates := map[int]candidate{}
		base := arrayOptionBase(procedure.file)
		procedureHasErrorHandling := arrayProcedureHasErrorHandling(proc)
		if participantCtx.arrayStats != nil {
			participantCtx.arrayStats.addCFGWalk()
		}
		baseView := proc.Graph.View(vbacfg.EdgeFilter{})
		walkArrayCFGWithStopStats(&baseView, procedure.file.Lines, arrayInitialState(procedure.variables), func(text string, line int, in arrayFlowState) arrayFlowState {
			if lhs, rhs, indexed, ok := arrayAssignment(text); ok && !indexed && strings.EqualFold(lhs, proc.Name) {
				value, known := arrayExpressionState(rhs, in, ctx)
				returnCandidates[line] = candidate{value: value, ok: known && value.kind == arrayAllocated && value.knownArray && value.origin != arrayOriginRangeValue}
			}
			out, _ := (Analyzer{}).arrayTransfer(procedure.file, proc, ctx, procedure.variables, in, text, line, procedure.constants, nil)
			return out
		}, func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState {
			return applyArrayAllocationGuard(out, block.Statement, edge, arrayAllocationGuards, procedure.variables)
		}, func(text string, _ int) bool {
			return !procedureHasErrorHandling && arraySummaryStatementAlwaysFails(text, base, procedure.constants)
		}, participantCtx.arrayStats)
		conditionalValue, hasConditional := arrayConditionalReturnSummary(procedure.file, proc)
		descriptorValue, hasDescriptor := arrayDescriptorArrayReturnSummary(procedure.file, proc, participantCtx)
		returnLines := make([]int, 0, len(returnCandidates))
		for line := range returnCandidates {
			returnLines = append(returnLines, line)
		}
		sort.Ints(returnLines)
		returns := make([]candidate, 0, len(returnLines))
		for _, line := range returnLines {
			returns = append(returns, returnCandidates[line])
		}
		if len(returns) == 0 || !proc.Graph.IsDefinitelyAssigned(proc.Graph.NormalExit, vbacfg.Variable{Scope: procedureir.ScopeLocal, Name: proc.Name}, vbacfg.EdgeFilter{NormalOnly: true}) {
			if hasConditional {
				return candidate{value: conditionalValue, ok: true}, true
			}
			if hasDescriptor {
				return candidate{value: descriptorValue, ok: true}, true
			}
			return candidate{}, false
		}
		valid := returns[0].ok
		value := returns[0].value
		for _, returned := range returns[1:] {
			if !returned.ok || !arrayReturnValueCompatible(proc, value, returned.value) {
				valid = false
				break
			}
			value = meetArrayValue(value, returned.value)
		}
		if !valid && hasConditional {
			return candidate{value: conditionalValue, ok: true}, true
		}
		if !valid && hasDescriptor {
			return candidate{value: descriptorValue, ok: true}, true
		}
		return candidate{value: value, ok: valid}, true
	}

	dependents := make(map[string][]int)
	for index, procedure := range procedures {
		for call := range procedure.proc.Calls.All() {
			resolution := call.Resolution
			if participantCtx.procedureResolver != nil {
				resolution = participantCtx.procedureResolver.ResolveCall(call)
			}
			for _, candidate := range resolution.Candidates {
				name := strings.ToLower(strings.TrimSpace(candidate.QualifiedName))
				if dot := strings.LastIndexByte(name, '.'); dot >= 0 {
					name = name[dot+1:]
				}
				if name != "" {
					dependents[name] = append(dependents[name], index)
				}
			}
		}
	}
	for name := range dependents {
		sort.Ints(dependents[name])
	}
	contributions := make(map[string]candidate, len(procedures))
	present := make(map[string]bool, len(procedures))
	groups := make(map[string]map[string]candidate)
	summaries := map[string]arrayValue{}
	qualifiedSummaries := map[string]arrayValue{}
	documentedSummaries := inferDocumentedArrayReturnSummaries(files)
	documentedBareSummaries := inferDocumentedNonEmptyArrayReturnSummaries(files)
	documentedLowerBoundSummaries := inferDocumentedArrayReturnLowerBoundSummaries(files)
	for key, value := range documentedSummaries {
		qualifiedSummaries[key] = value
	}
	queue := make([]int, len(procedures))
	queued := make([]bool, len(procedures))
	for index := range procedures {
		queue[index] = index
		queued[index] = true
	}
	for head := 0; head < len(queue); head++ {
		index := queue[head]
		queued[index] = false
		procedure := procedures[index]
		name := strings.ToLower(strings.TrimSpace(procedure.proc.Name))
		key := arrayProcedureKey(procedure.proc)
		value, hasContribution := evaluate(procedure, summaries)
		if hasContribution && value.ok {
			qualifiedSummaries[key] = value.value
		} else {
			delete(qualifiedSummaries, key)
		}
		if ambiguousReturnNames[name] {
			// Summary lookups use bare names for compatibility with the existing
			// expression resolver. A duplicate bare name is therefore permanently
			// ambiguous for this revision. Keep it at the unknown bottom of the
			// lattice instead of allowing iteration order to delete and recreate a
			// summary while duplicate candidates are evaluated.
			continue
		}
		if head >= len(procedures) && participantCtx.arrayStats != nil {
			participantCtx.arrayStats.addRevisit()
		}
		if present[key] == hasContribution && (!hasContribution || arrayValueEqual(contributions[key].value, value.value) && contributions[key].ok == value.ok) {
			continue
		}
		if hasContribution {
			contributions[key] = value
			present[key] = true
			if groups[name] == nil {
				groups[name] = map[string]candidate{}
			}
			groups[name][key] = value
		} else {
			delete(contributions, key)
			delete(present, key)
			if group := groups[name]; group != nil {
				delete(group, key)
				if len(group) == 0 {
					delete(groups, name)
				}
			}
		}
		old, hadOld := summaries[name]
		group := groups[name]
		if len(group) == 1 {
			for _, value := range group {
				if value.ok {
					summaries[name] = value.value
				} else {
					delete(summaries, name)
				}
			}
		} else {
			delete(summaries, name)
		}
		fresh, hasFresh := summaries[name]
		changed := hadOld != hasFresh || (hasFresh && !arrayValueEqual(old, fresh))
		if !changed {
			continue
		}
		for _, dependent := range dependents[name] {
			if !queued[dependent] {
				queued[dependent] = true
				queue = append(queue, dependent)
			}
		}
	}
	for key, value := range documentedSummaries {
		if _, exists := qualifiedSummaries[key]; !exists {
			qualifiedSummaries[key] = value
		}
	}
	for key, value := range documentedBareSummaries {
		name := key
		if dot := strings.LastIndexByte(name, '.'); dot >= 0 {
			name = name[dot+1:]
		}
		if !ambiguousReturnNames[name] {
			if _, exists := summaries[name]; !exists {
				summaries[name] = value
			}
		}
	}
	for key, value := range documentedLowerBoundSummaries {
		name := key
		if dot := strings.LastIndexByte(name, '.'); dot >= 0 {
			name = name[dot+1:]
		}
		if !ambiguousReturnNames[name] {
			// The lower-bound-only contract is deliberately more precise than
			// an inferred allocated summary: it preserves the possible empty
			// input path while still proving the loop's lower-side index bound.
			// Prefer it for the bare lookup so a generic fixed-point result
			// cannot erase the retained UBound diagnostic.
			summaries[name] = value
		}
	}
	return arrayReturnSummarySet{bare: summaries, qualified: qualifiedSummaries}
}

func arrayProcedureHasErrorHandling(proc sourceProcedure) bool {
	for statement := range proc.Statements.All() {
		if strings.HasPrefix(strings.ToLower(strings.TrimSpace(statement.Text)), "on error ") {
			return true
		}
	}
	return false
}

func arraySummaryStatementAlwaysFails(text string, base int, constants map[string]int) bool {
	match := arrayRedimRe.FindStringSubmatch(text)
	if len(match) == 0 {
		return false
	}
	for _, clause := range splitArgs(match[2]) {
		redim, direct := parseDirectArrayRedimClause(clause)
		if !direct {
			continue
		}
		if impossibleArrayBounds(parseArrayDimensionsWithConstants(redim.dimensions, base, constants)) {
			return true
		}
	}
	return false
}

func arrayValueEqual(left, right arrayValue) bool {
	return arrayValueCompatible(left, right) && left.mayBeEmpty == right.mayBeEmpty
}

// arrayReturnValueCompatible keeps a return summary when every normal return
// path produces an array with the same allocation provenance. Return branches
// may legitimately produce different shapes (for example Array() for an
// empty result and ReDim 1 To length for a non-empty result); that shape
// mismatch must not erase the stronger fact that the returned value is an
// array. Callers still receive no shape bounds from such a summary.
func arrayReturnValueCompatible(proc sourceProcedure, left, right arrayValue) bool {
	if arrayValueCompatible(left, right) {
		return true
	}
	if proc.ProcedureKind != procedureir.ProcedurePropertyGet {
		return false
	}
	return left.kind == arrayAllocated && right.kind == arrayAllocated && left.knownArray && right.knownArray && left.origin == right.origin && left.allocationProbe == right.allocationProbe && left.safeBoundProbe == right.safeBoundProbe && left.allocationCountSource == right.allocationCountSource && left.returnNonEmptyArrayParameter == right.returnNonEmptyArrayParameter && left.returnPositiveScalarParameter == right.returnPositiveScalarParameter && left.nonEmptySource == right.nonEmptySource && left.returnDescriptorSourceParameter == right.returnDescriptorSourceParameter && left.returnDescriptorStartParameter == right.returnDescriptorStartParameter && left.returnDescriptorLengthParameter == right.returnDescriptorLengthParameter && left.returnDescriptorLowerParameter == right.returnDescriptorLowerParameter
}

func arrayValueCompatible(left, right arrayValue) bool {
	return left.kind == right.kind && left.knownArray == right.knownArray && left.origin == right.origin && left.allocationProbe == right.allocationProbe && left.allocationCountSource == right.allocationCountSource && left.returnNonEmptyArrayParameter == right.returnNonEmptyArrayParameter && left.returnPositiveScalarParameter == right.returnPositiveScalarParameter && left.nonEmptySource == right.nonEmptySource && left.returnDescriptorSourceParameter == right.returnDescriptorSourceParameter && left.returnDescriptorStartParameter == right.returnDescriptorStartParameter && left.returnDescriptorLengthParameter == right.returnDescriptorLengthParameter && left.returnDescriptorLowerParameter == right.returnDescriptorLowerParameter && arrayDimensionsEqual(left.dimensions, right.dimensions) && arrayDimensionsEqual(left.preserveShape, right.preserveShape)
}
