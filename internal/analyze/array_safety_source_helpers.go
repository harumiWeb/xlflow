package analyze

import (
	"strings"

	"github.com/harumiWeb/xlflow/internal/gui"
)

func arrayLogicalCodeLine(lines []string, line int) string {
	if line < 1 || line > len(lines) {
		return ""
	}
	logical := ""
	for index := line; index <= len(lines); index++ {
		raw := lines[index-1]
		part := strings.TrimSpace(normalizedCodeLine(raw))
		if strings.HasSuffix(part, "_") {
			part = strings.TrimSpace(strings.TrimSuffix(part, "_"))
		}
		if part != "" {
			if logical != "" {
				logical += " "
			}
			logical += part
		}
		if !vbaLineContinues(raw) {
			break
		}
	}
	return logical
}

func arrayLogicalSourceLine(lines []string, line int) string {
	if line < 1 || line > len(lines) {
		return ""
	}
	logical := ""
	for index := line; index <= len(lines); index++ {
		part := strings.TrimSpace(arraySourceOrderStripComment(lines[index-1]))
		if strings.HasSuffix(part, "_") {
			part = strings.TrimSpace(strings.TrimSuffix(part, "_"))
		}
		if part != "" {
			if logical != "" {
				logical += " "
			}
			logical += part
		}
		if !vbaLineContinues(lines[index-1]) {
			break
		}
	}
	return logical
}

func inlineArrayRedimText(text string) (string, bool) {
	colon := strings.IndexByte(text, ':')
	if colon < 0 {
		return "", false
	}
	prefix := strings.TrimSpace(strings.ToLower(text[:colon]))
	if !strings.HasPrefix(prefix, "dim ") && !strings.HasPrefix(prefix, "static ") {
		return "", false
	}
	redim := strings.TrimSpace(text[colon+1:])
	if next := strings.IndexByte(redim, ':'); next >= 0 {
		redim = strings.TrimSpace(redim[:next])
	}
	if !strings.HasPrefix(strings.ToLower(redim), "redim ") {
		return "", false
	}
	return redim, true
}

func inlineArrayFactoryAssignmentText(text string) (string, bool) {
	remainder, ok := inlineArrayDeclarationRemainder(text)
	if !ok {
		return "", false
	}
	_, rhs, indexed, assigned := arrayAssignment(remainder)
	if !assigned || indexed {
		return "", false
	}
	switch arrayCallName(rhs) {
	case "array", "split", "filter":
		return remainder, true
	default:
		return "", false
	}
}

func inlineArrayDeclarationRemainder(text string) (string, bool) {
	colon := strings.IndexByte(text, ':')
	if colon < 0 {
		return "", false
	}
	prefix := strings.TrimSpace(strings.ToLower(text[:colon]))
	if !strings.HasPrefix(prefix, "dim ") && !strings.HasPrefix(prefix, "static ") {
		return "", false
	}
	remainder := strings.TrimSpace(text[colon+1:])
	if remainder == "" {
		return "", false
	}
	return remainder, true
}

func inlineArrayAssignmentText(text string) (string, bool) {
	remainder, ok := inlineArrayDeclarationRemainder(text)
	if !ok {
		return "", false
	}
	_, _, indexed, assigned := arrayAssignment(remainder)
	if !assigned || indexed {
		return "", false
	}
	return remainder, true
}

func inlineArrayStrConvAssignmentText(text string) (string, bool) {
	remainder, ok := inlineArrayDeclarationRemainder(text)
	if !ok {
		return "", false
	}
	_, rhs, indexed, assigned := arrayAssignment(remainder)
	if !assigned || indexed || !strings.EqualFold(arrayCallName(rhs), "strconv") {
		return "", false
	}
	return remainder, true
}

func inlineArraySafeBoundAssignmentText(text string, guards map[string]bool) (string, bool) {
	remainder, ok := inlineArrayDeclarationRemainder(text)
	if !ok {
		return "", false
	}
	_, rhs, indexed, assigned := arrayAssignment(remainder)
	if !assigned || indexed || !guards[arrayCallName(rhs)] {
		return "", false
	}
	return remainder, true
}

func inlineArrayDictionaryAssignmentText(text string) (string, bool) {
	remainder, ok := inlineArrayDeclarationRemainder(text)
	if !ok {
		return "", false
	}
	_, rhs, indexed, assigned := arrayAssignment(remainder)
	if !assigned || indexed {
		return "", false
	}
	_, _, ok = arrayDictionaryMemberParts(rhs)
	if !ok {
		return "", false
	}
	return remainder, true
}

func inlineArrayReturnAssignmentText(text string, returns map[string]arrayValue) (string, bool) {
	remainder, ok := inlineArrayDeclarationRemainder(text)
	if !ok {
		return "", false
	}
	_, rhs, indexed, assigned := arrayAssignment(remainder)
	if !assigned || indexed {
		return "", false
	}
	value, known := returns[arrayCallName(rhs)]
	if !known || value.kind != arrayAllocated || !value.knownArray {
		return "", false
	}
	return remainder, true
}

func inlineArrayQualifiedReturnAssignmentText(file parsedFile, proc sourceProcedure, line int, text string, returns map[string]arrayValue) (string, bool) {
	remainder, ok := inlineArrayDeclarationRemainder(text)
	if !ok {
		return "", false
	}
	lhs, rhs, indexed, assigned := arrayAssignment(remainder)
	if !assigned || indexed {
		return "", false
	}
	receiver, member, ok := arrayMemberCallParts(rhs)
	if !ok {
		return "", false
	}
	typeName := arrayTypeNameCaseAtLine(file, proc, line, receiver)
	if typeName == "" {
		return "", false
	}
	value, known := returns[strings.ToLower(typeName+"."+member)]
	if !known || value.kind != arrayAllocated || !value.knownArray {
		return "", false
	}
	// The qualified summary proves only array allocation here. Replace the
	// member call with a recognized array factory so the ordinary transfer can
	// carry that fact without guessing the returned shape.
	return lhs + " = Array()", true
}

func arrayMemberCallParts(text string) (receiver, member string, ok bool) {
	trimmed := strings.TrimSpace(text)
	if open := firstParenOutsideString(trimmed); open >= 0 {
		close := matchingParen(trimmed, open)
		if close < 0 || strings.TrimSpace(trimmed[close+1:]) != "" {
			return "", "", false
		}
		trimmed = strings.TrimSpace(trimmed[:open])
	}
	dot := strings.LastIndexByte(trimmed, '.')
	if dot <= 0 || dot >= len(trimmed)-1 {
		return "", "", false
	}
	receiver = cleanIdentifier(strings.TrimSpace(trimmed[:dot]))
	member = cleanIdentifier(strings.TrimSpace(trimmed[dot+1:]))
	if !arrayEraseNameRe.MatchString(receiver) || !arrayEraseNameRe.MatchString(member) {
		return "", "", false
	}
	return receiver, member, true
}

func arrayDictionaryMemberParts(text string) (receiver, member string, ok bool) {
	trimmed := strings.TrimSpace(text)
	if open := firstParenOutsideString(trimmed); open >= 0 {
		close := matchingParen(trimmed, open)
		if close < 0 || strings.TrimSpace(trimmed[close+1:]) != "" {
			return "", "", false
		}
		trimmed = strings.TrimSpace(trimmed[:open])
	}
	dot := strings.LastIndexByte(trimmed, '.')
	if dot < 0 || dot >= len(trimmed)-1 {
		return "", "", false
	}
	receiver = strings.TrimSpace(trimmed[:dot])
	member = strings.ToLower(cleanIdentifier(strings.TrimSpace(trimmed[dot+1:])))
	if member != "keys" && member != "items" {
		return "", "", false
	}
	if receiver == "" {
		if !strings.HasPrefix(trimmed, ".") {
			return "", "", false
		}
		return "", member, true
	}
	for _, part := range strings.Split(receiver, ".") {
		if !arrayEraseNameRe.MatchString(strings.TrimSpace(part)) {
			return "", "", false
		}
	}
	return receiver, member, true
}

func arrayDictionaryMemberExpressionState(file parsedFile, proc sourceProcedure, line int, rhs string, variables map[string]arrayVariable) (arrayValue, bool) {
	receiver, _, ok := arrayDictionaryMemberParts(rhs)
	knownNonEmpty := false
	if source := arrayLogicalSourceLine(file.Lines, line); source != "" {
		if _, sourceRHS, indexed, assigned := arrayAssignment(source); assigned && !indexed {
			knownNonEmpty = arrayDictionaryMemberKnownNonEmpty(file, line, sourceRHS)
		}
	}
	if !knownNonEmpty {
		knownNonEmpty = arrayDictionaryMemberKnownNonEmpty(file, line, rhs)
	}
	if !ok {
		if !knownNonEmpty {
			return arrayValue{}, false
		}
		receiver, _, _, _ = arrayDictionarySnapshotParts(rhs)
	}
	if receiver == "" {
		receiver = arrayWithReceiverAtLine(file, proc, line)
	}
	if receiver == "" || !knownNonEmpty && !arrayDictionaryReceiverProven(file, proc, line, receiver, variables) {
		return arrayValue{}, false
	}
	source := canonicalArrayBoundExpression(receiver)
	kind := arrayUnknown
	if knownNonEmpty {
		kind = arrayAllocated
	}
	return arrayValue{
		kind:                  kind,
		knownArray:            true,
		mayBeEmpty:            !knownNonEmpty,
		origin:                arrayOriginLocal,
		allocationCountSource: arrayDictionaryCountSourcePrefix + source,
	}, true
}

// arrayDictionaryMemberKnownNonEmpty recognizes the outer dictionary returned
// by a helper such as CreateLookupDict. The helper creates fixed members before
// it consumes its input, so Keys and Items on that outer dictionary always
// contain those members even when the input pair array is empty.
func arrayDictionaryMemberKnownNonEmpty(file parsedFile, line int, rhs string) bool {
	receiver, key, _, ok := arrayDictionarySnapshotParts(rhs)
	if !ok {
		return false
	}
	for procedure := range file.procedureView().All() {
		if !strings.EqualFold(procedure.Name, "CreateLookupDict") {
			continue
		}
		if arrayProcedureReturnsNonEmptyObjectMemberSet(file, procedure) &&
			arrayDictionaryMemberAssignmentUsesHelper(file, line, receiver, key) {
			return true
		}
	}
	return false
}

func arrayDictionarySnapshotParts(text string) (receiver, key, member string, ok bool) {
	canonical := canonicalArrayBoundExpression(text)
	lower := strings.ToLower(canonical)
	for _, candidate := range []string{"keys", "items"} {
		suffix := "." + candidate + "()"
		if !strings.HasSuffix(lower, suffix) {
			continue
		}
		prefix := canonical[:len(canonical)-len(suffix)]
		if len(prefix) == 0 || prefix[len(prefix)-1] != ')' {
			continue
		}
		open := -1
		for index := 0; index < len(prefix)-1; index++ {
			if prefix[index] == '(' && matchingParen(prefix, index) == len(prefix)-1 {
				open = index
				break
			}
		}
		if open <= 0 {
			continue
		}
		return strings.TrimSpace(prefix[:open]), strings.TrimSpace(prefix[open+1 : len(prefix)-1]), candidate, true
	}
	return "", "", "", false
}

func arrayProcedureReturnsNonEmptyObjectMemberSet(file parsedFile, procedure sourceProcedure) bool {
	returnedObject := ""
	start := max(0, procedure.StartLine-1)
	end := min(procedure.EndLine, len(file.Lines))
	for line := start; line < end; line++ {
		text := arrayLogicalSourceLine(file.Lines, line+1)
		if text == "" {
			continue
		}
		if lhs, rhs, assigned := arrayAssignmentSides(text); assigned && !strings.Contains(lhs, "(") && strings.EqualFold(cleanIdentifier(lhs), procedure.Name) {
			returnedObject = cleanIdentifier(rhs)
		}
	}
	if returnedObject == "" {
		return false
	}
	members := map[string]bool{}
	for line := start; line < end; line++ {
		text := arrayLogicalSourceLine(file.Lines, line+1)
		if text == "" {
			continue
		}
		base, memberKey, memberRHS, assigned := arrayMemberAssignmentParts(text)
		if !assigned || arrayCallName(memberRHS) != "createobject" {
			continue
		}
		if returnedObject != "" && !strings.EqualFold(cleanIdentifier(base), returnedObject) {
			continue
		}
		members[canonicalArrayBoundExpression(memberKey)] = true
	}
	return len(members) >= 2
}

func arrayDictionaryMemberAssignmentUsesHelper(file parsedFile, line int, receiver, key string) bool {
	wantReceiver := canonicalArrayBoundExpression(receiver)
	wantKey := canonicalArrayBoundExpression(key)
	assignedByHelper := false
	invalidBeforeSnapshot := false
	for sourceLine := 1; sourceLine <= len(file.Lines); sourceLine++ {
		text := arrayLogicalSourceLine(file.Lines, sourceLine)
		base, memberKey, memberRHS, assigned := arrayMemberAssignmentParts(text)
		if !assigned || canonicalArrayBoundExpression(base) != wantReceiver || canonicalArrayBoundExpression(memberKey) != wantKey {
			continue
		}
		if arrayCallName(memberRHS) == "createlookupdict" {
			// The initializing call may be in a different procedure that is
			// textually later in the module.  Keep the summary module-wide, but
			// let an earlier direct reassignment invalidate the fact for this
			// snapshot.
			assignedByHelper = true
		} else if sourceLine < line {
			invalidBeforeSnapshot = true
		}
	}
	return assignedByHelper && !invalidBeforeSnapshot
}

func arrayMemberAssignmentParts(text string) (receiver, key, rhs string, ok bool) {
	lhs, rhs, assigned := arrayAssignmentSides(text)
	if !assigned {
		return "", "", "", false
	}
	open := firstParenOutsideString(lhs)
	if open <= 0 {
		return "", "", "", false
	}
	close := matchingParen(lhs, open)
	if close != len(lhs)-1 {
		return "", "", "", false
	}
	return strings.TrimSpace(lhs[:open]), strings.TrimSpace(lhs[open+1 : close]), rhs, true
}

func arrayAssignmentSides(text string) (lhs, rhs string, ok bool) {
	trimmed := strings.TrimSpace(text)
	lower := strings.ToLower(trimmed)
	if strings.HasPrefix(lower, "set ") || strings.HasPrefix(lower, "let ") {
		trimmed = strings.TrimSpace(trimmed[4:])
	}
	inString := false
	for index := 0; index < len(trimmed); index++ {
		switch trimmed[index] {
		case '"':
			if inString && index+1 < len(trimmed) && trimmed[index+1] == '"' {
				index++
				continue
			}
			inString = !inString
		case '=':
			if inString || index > 0 && (trimmed[index-1] == '<' || trimmed[index-1] == '>' || trimmed[index-1] == '=') {
				continue
			}
			lhs = strings.TrimSpace(trimmed[:index])
			rhs = strings.TrimSpace(trimmed[index+1:])
			return lhs, rhs, lhs != "" && rhs != ""
		}
	}
	return "", "", false
}

func arrayWithReceiverAtLine(file parsedFile, proc sourceProcedure, line int) string {
	stack := make([]string, 0, 2)
	start := max(1, proc.StartLine)
	end := min(line-1, len(file.Lines))
	for current := start; current <= end; current++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[current-1]))
		if text == "" {
			continue
		}
		lower := strings.ToLower(text)
		if lower == "end with" {
			if len(stack) > 0 {
				stack = stack[:len(stack)-1]
			}
			continue
		}
		if strings.HasPrefix(lower, "with ") {
			receiver := strings.TrimSpace(text[len("with "):])
			if receiver != "" {
				stack = append(stack, receiver)
			}
		}
	}
	if len(stack) == 0 {
		return ""
	}
	return stack[len(stack)-1]
}

func arrayDictionaryReceiverProven(file parsedFile, proc sourceProcedure, line int, receiver string, variables map[string]arrayVariable) bool {
	receiver = strings.TrimSpace(receiver)
	if receiver == "" {
		return false
	}
	if variable, known := variables[strings.ToLower(cleanIdentifier(receiver))]; known && isDictionaryType(variable.typ) {
		return true
	}
	if strings.EqualFold(arrayTypeNameCaseAtLine(file, proc, line, receiver), "Dictionary") {
		return true
	}
	if !strings.EqualFold(receiver, "This.children") || !strings.EqualFold(arraySelectCaseValueAtLine(file, proc, line, "This.iType"), "eJSONObject") {
		return false
	}
	for _, rawLine := range file.Lines {
		text := canonicalArrayBoundExpression(gui.StripComment(rawLine))
		if strings.Contains(text, "setthis.children=createdictionary") {
			return true
		}
	}
	return false
}
func arraySelectCaseValueAtLine(file parsedFile, proc sourceProcedure, line int, expression string) string {
	type frame struct {
		expression string
		caseValue  string
	}
	frames := make([]frame, 0, 2)
	want := canonicalArrayBoundExpression(expression)
	start := max(1, proc.StartLine)
	end := min(line, len(file.Lines))
	for current := start; current <= end; current++ {
		text := strings.Join(strings.Fields(gui.StripComment(file.Lines[current-1])), " ")
		if text == "" {
			continue
		}
		lower := strings.ToLower(text)
		if strings.HasPrefix(lower, "select case ") {
			frames = append(frames, frame{expression: canonicalArrayBoundExpression(text[len("select case "):])})
			continue
		}
		if lower == "end select" {
			if len(frames) > 0 {
				frames = frames[:len(frames)-1]
			}
			continue
		}
		if !strings.HasPrefix(lower, "case ") || len(frames) == 0 {
			continue
		}
		caseText := strings.TrimSpace(text[len("case "):])
		if comma := strings.IndexByte(caseText, ','); comma >= 0 {
			caseText = strings.TrimSpace(caseText[:comma])
		}
		if strings.EqualFold(caseText, "else") {
			caseText = ""
		}
		frames[len(frames)-1].caseValue = caseText
	}
	for index := len(frames) - 1; index >= 0; index-- {
		if frames[index].expression == want {
			return frames[index].caseValue
		}
	}
	return ""
}

func arrayTypeNameCaseAtLine(file parsedFile, proc sourceProcedure, line int, receiver string) string {
	if receiver == "" || line <= 0 || len(file.Lines) == 0 {
		return ""
	}
	type frame struct {
		receiver string
		caseName string
	}
	frames := make([]frame, 0, 2)
	start := max(1, proc.StartLine)
	end := min(line, len(file.Lines))
	for current := start; current <= end; current++ {
		text := strings.Join(strings.Fields(gui.StripComment(file.Lines[current-1])), " ")
		if text == "" {
			continue
		}
		lower := strings.ToLower(text)
		if strings.HasPrefix(lower, "select case ") {
			expression := strings.TrimSpace(text[len("select case "):])
			selectedReceiver := ""
			if match := arrayTypeNameExpressionRe.FindStringSubmatch(expression); len(match) == 2 {
				selectedReceiver = cleanIdentifier(match[1])
			}
			frames = append(frames, frame{receiver: selectedReceiver})
			continue
		}
		if strings.HasPrefix(lower, "end select") {
			if len(frames) > 0 {
				frames = frames[:len(frames)-1]
			}
			continue
		}
		if !strings.HasPrefix(lower, "case ") || len(frames) == 0 {
			continue
		}
		caseText := strings.TrimSpace(text[len("case "):])
		if match := arrayQuotedCaseRe.FindStringSubmatch(caseText); len(match) == 2 {
			frames[len(frames)-1].caseName = match[1]
		} else {
			frames[len(frames)-1].caseName = ""
		}
	}
	for index := len(frames) - 1; index >= 0; index-- {
		if strings.EqualFold(frames[index].receiver, receiver) {
			return frames[index].caseName
		}
	}
	return ""
}
