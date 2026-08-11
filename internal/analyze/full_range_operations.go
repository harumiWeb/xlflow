package analyze

import (
	"regexp"
	"strconv"
	"strings"
	"unicode"

	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// VBA242 deliberately starts from the syntactic range shape and only uses the
// Excel type database to validate an explicit receiver.  This keeps late-bound
// and recovered expressions out of the finding set while still recognizing
// the unqualified Excel built-ins.
type vba242Shape struct {
	kind       string
	expression string
	start      int
	end        int
}

type vba242Call struct {
	name      string
	start     int
	open      int
	close     int
	arguments string
}

var vba242ReceiverRe = regexp.MustCompile(`(?i)([A-Za-z_][A-Za-z0-9_]*(?:\s*\([^()\r\n]*\))?(?:\s*\.\s*[A-Za-z_][A-Za-z0-9_]*(?:\s*\([^()\r\n]*\))?)*)\s*$`)
var vba242MemberChainRe = regexp.MustCompile(`^(?:\s*(?:_\s*)?\.\s*[A-Za-z_][A-Za-z0-9_]*)*\s*(?:_\s*)?\.\s*$`)

var vba242SinkMethods = map[string]bool{
	"calculate":     true,
	"find":          true,
	"replace":       true,
	"sort":          true,
	"setrange":      true,
	"clear":         true,
	"clearcontents": true,
	"clearformats":  true,
	"delete":        true,
	"copy":          true,
	"pastespecial":  true,
}

var vba242SinkProperties = map[string]bool{
	"value": true, "value2": true,
	"formula": true, "formula2": true, "formular1c1": true, "formula2r1c1": true, "formular1c1local": true, "formula2r1c1local": true,
	"numberformat": true, "numberformatlocal": true,
	"horizontalalignment": true, "verticalalignment": true,
	"wraptext": true, "orientation": true, "indentlevel": true,
	"font": true, "bold": true, "italic": true, "underline": true,
	"color": true, "colorindex": true, "pattern": true,
	"borders": true, "border": true,
	"style": true, "locked": true, "hidden": true, "shrinktofit": true,
	"mergecells": true, "readingorder": true,
	"rowheight": true, "columnwidth": true, "autofit": true,
	"formatconditions": true, "validation": true,
}

// expensiveFullRangeOperationFindings reports only operations on a full range
// shape. Merely obtaining a Range/UsedRange object is intentionally harmless;
// the diagnostic is tied to a mutating or repeatedly expensive sink.
func (a Analyzer) expensiveFullRangeOperationFindings(file parsedFile, proc sourceProcedure) []Finding {
	if !a.Config.Analyze.DetectExpensiveFullRangeOperations {
		return nil
	}
	regions := excelLoopRegions(proc)
	var findings []Finding
	seen := map[int]int{}
	for _, statement := range proc.Statements {
		if statement.Recovered {
			continue
		}
		isLoopHeader := isExcelLoopKind(statement.Kind)
		code := statement.Text
		if line := statement.Range.StartLine; line > 0 && line <= len(file.Lines) {
			// The IR represents an unqualified member call such as
			// `Rows("1:1").Calculate` as a `.Calculate` member statement. Use
			// the source line when it carries the range receiver.
			lineCode := strings.TrimSpace(file.Lines[line-1])
			if vba242LooksLikeRangeOperation(lineCode) {
				if isLoopHeader {
					// Loop headers can carry a condition that is evaluated on
					// every iteration. Their IR statement text may include child
					// body expressions, so use only the source header here.
					code = lineCode
				} else {
					code = lineCode + " " + strings.TrimSpace(statement.Text)
				}
			} else if isLoopHeader {
				// Do not mine body expressions from a loop-header IR node.
				continue
			}
		}
		code = vba242Code(code)
		if strings.TrimSpace(code) == "" {
			continue
		}
		shapes := a.vba242Shapes(file, proc, code, statement.Range.StartLine-1)
		shapes = vba242EffectiveShapes(code, shapes)
		if len(shapes) == 0 || !vba242HasExpensiveSink(code, shapes) {
			continue
		}
		inLoop := false
		if statement.SyntaxKind == "do_condition" {
			parent := procedureStatementByID(proc, statement.ParentID)
			inLoop = parent.Kind == procedureir.StatementDo
		}
		for _, region := range containingExcelLoops(regions, statement.ID, statement.Range.StartLine) {
			// containingExcelLoops also has a line-based fallback for malformed
			// source. Only a CFG-reachable body is a repeated operation.
			if region.StatementID == statement.ID || region.Body[statement.ID] {
				inLoop = true
				break
			}
		}
		severity := "information"
		if inLoop {
			severity = "warning"
		}
		if previous, ok := seen[statement.Range.StartLine]; ok {
			if inLoop && findings[previous].Severity == "information" {
				findings[previous].Severity = "warning"
				findings[previous].Message += " It executes inside a reachable loop."
				findings[previous].Reason += " Repeating the operation in a loop multiplies the range traversal and COM cost."
			}
			continue
		}
		seen[statement.Range.StartLine] = len(findings)
		kinds := make([]string, 0, len(shapes))
		for _, shape := range shapes {
			if !containsStringFold(kinds, shape.kind) {
				kinds = append(kinds, shape.kind)
			}
		}
		message := "Expensive operation targets an entire " + strings.Join(kinds, "/") + " range."
		if inLoop {
			message += " It executes inside a reachable loop."
		}
		reason := "Whole-row, whole-column, and whole-sheet operations make Excel inspect far more cells than the data operation normally requires."
		if inLoop {
			reason += " Repeating the operation in a loop multiplies the range traversal and COM cost."
		}
		suggestion := "Derive a bounded last row or last column and apply the operation only to that range. Keep intentional one-time formatting suppressed locally when the full range is required."
		findings = append(findings, a.simpleFinding(file, proc, statement.Range.StartLine, "VBA242", severity, message, reason, suggestion))
	}
	return findings
}

func containsStringFold(values []string, want string) bool {
	for _, value := range values {
		if strings.EqualFold(value, want) {
			return true
		}
	}
	return false
}

func vba242LooksLikeRangeOperation(code string) bool {
	lower := strings.ToLower(code)
	for _, name := range []string{"columns", "rows", "range", "cells", "usedrange", "entirerow", "entirecolumn"} {
		if strings.Contains(lower, name) {
			return true
		}
	}
	return false
}

func (a Analyzer) vba242Shapes(file parsedFile, proc sourceProcedure, code string, line int) []vba242Shape {
	var out []vba242Shape
	shadowed := vba242ShadowedRoots(file, proc)
	calls := vba242ScanCalls(code, map[string]bool{"columns": true, "rows": true, "range": true, "cells": true})
	for _, call := range calls {
		receiver := vba242Receiver(code, call.start)
		kind := ""
		switch strings.ToLower(call.name) {
		case "columns":
			if vba242FullColumnArgument(call.arguments) {
				kind = "column"
			}
		case "rows":
			if vba242FullRowArgument(call.arguments) {
				kind = "row"
			}
		case "range":
			if vba242FullRangeArgument(call.arguments) {
				kind = vba242RangeKind(call.arguments)
			} else if vba242MaxCellsRangeArgument(call.arguments) {
				kind = "sheet"
			}
		case "cells":
			if strings.TrimSpace(call.arguments) == "" {
				kind = "sheet"
			}
		}
		if kind == "" || !a.vba242ReceiverAllowed(file, receiver, call.name, kind, line) {
			continue
		}
		if receiver == "" && shadowed[strings.ToLower(call.name)] {
			continue
		}
		out = append(out, vba242Shape{kind: kind, expression: strings.TrimSpace(code[call.start : call.close+1]), start: call.start, end: call.close + 1})
	}

	// Cells is also a VBA property with no parentheses (for example,
	// `ws.Cells.Formula = ...`). Scan identifier tokens outside strings and
	// comments; call forms were already handled above.
	for _, token := range vba242Identifiers(code) {
		lower := strings.ToLower(token.name)
		if lower != "rows" && lower != "columns" && lower != "cells" && lower != "usedrange" && lower != "entirerow" && lower != "entirecolumn" {
			continue
		}
		if (token.name == "Cells" || token.name == "Rows" || token.name == "Columns") && token.nextNonSpace == '(' {
			continue
		}
		receiver := vba242Receiver(code, token.start)
		kind := map[string]string{"rows": "row", "columns": "column", "cells": "sheet", "usedrange": "used", "entirerow": "row", "entirecolumn": "column"}[lower]
		if !a.vba242ReceiverAllowed(file, receiver, token.name, kind, line) {
			continue
		}
		if receiver == "" && shadowed[lower] {
			continue
		}
		out = append(out, vba242Shape{kind: kind, expression: strings.TrimSpace(code[token.start:token.end]), start: token.start, end: token.end})
	}
	return vba242DeduplicateShapes(out)
}

func vba242ShadowedRoots(file parsedFile, proc sourceProcedure) map[string]bool {
	shadowed := map[string]bool{}
	for _, declaration := range proc.Declarations {
		name := strings.ToLower(strings.TrimSpace(declaration.Name))
		if vba242IsRootName(name) {
			shadowed[name] = true
		}
	}
	for _, declaration := range file.IR.Declarations {
		name := strings.ToLower(strings.TrimSpace(declaration.Name))
		if vba242IsRootName(name) {
			shadowed[name] = true
		}
	}
	for _, procedure := range file.IR.Procedures {
		name := strings.ToLower(strings.TrimSpace(procedure.Symbol.Name))
		if vba242IsRootName(name) {
			shadowed[name] = true
		}
	}
	for _, access := range proc.Accesses {
		if access.Scope != procedureir.ScopeLocal && access.Scope != procedureir.ScopeParameter && access.Scope != procedureir.ScopeModule && access.Scope != procedureir.ScopeProject {
			continue
		}
		name := strings.ToLower(strings.TrimSpace(access.Name))
		if vba242IsRootName(name) {
			shadowed[name] = true
		}
	}
	for _, parameter := range proc.Params {
		name := strings.ToLower(strings.TrimSpace(parameter.Name))
		if vba242IsRootName(name) {
			shadowed[name] = true
		}
	}
	return shadowed
}

func vba242IsRootName(name string) bool {
	return name == "columns" || name == "rows" || name == "range" || name == "cells" || name == "usedrange"
}

func vba242DeduplicateShapes(shapes []vba242Shape) []vba242Shape {
	seen := map[string]bool{}
	out := make([]vba242Shape, 0, len(shapes))
	for _, shape := range shapes {
		key := strconv.Itoa(shape.start) + ":" + strconv.Itoa(shape.end) + ":" + shape.kind
		if seen[key] {
			continue
		}
		seen[key] = true
		out = append(out, shape)
	}
	return out
}

func (a Analyzer) vba242ReceiverAllowed(file parsedFile, receiver, member, kind string, line int) bool {
	receiver = strings.TrimSpace(receiver)
	if receiver == "" {
		// These names are Excel's unqualified built-ins. EntireRow and
		// EntireColumn require a proven Range receiver below, but UsedRange is
		// a well-known active-sheet property just like Rows/Columns.
		return kind != "row" && kind != "column" || strings.EqualFold(member, "Rows") || strings.EqualFold(member, "Columns") || strings.EqualFold(member, "Range") || strings.EqualFold(member, "UsedRange")
	}
	typ, resolved := resolveExcelExpressionType(file, a.typeDB, receiver, line, a.RootDir, a.Config)
	if !resolved {
		return false
	}
	lower := strings.ToLower(typ)
	if kind == "used" {
		return strings.Contains(lower, "worksheet")
	}
	if kind == "row" || kind == "column" {
		if strings.EqualFold(member, "EntireRow") || strings.EqualFold(member, "EntireColumn") {
			return isExcelRangeType(typ)
		}
	}
	// A Range.Columns/Rows call is bounded by that range; only worksheet,
	// workbook, or application receivers describe a sheet-wide shape.
	if isExcelRangeType(typ) {
		return false
	}
	return strings.Contains(lower, "worksheet") || strings.Contains(lower, "workbook") || strings.Contains(lower, "application")
}

func vba242FullColumnArgument(arguments string) bool {
	if strings.TrimSpace(arguments) == "" {
		return true
	}
	argument := vba242QuotedArgument(arguments)
	if argument == "" {
		value := strings.TrimSpace(arguments)
		number, err := strconv.Atoi(value)
		return err == nil && number > 0
	}
	argument = strings.ReplaceAll(strings.ToUpper(argument), "$", "")
	argument = strings.ReplaceAll(argument, " ", "")
	parts := strings.Split(argument, ":")
	if len(parts) > 2 || len(parts) == 0 {
		return false
	}
	for _, part := range parts {
		if part == "" || !vba242ColumnName(part) {
			return false
		}
	}
	return true
}

func vba242FullRowArgument(arguments string) bool {
	if strings.TrimSpace(arguments) == "" {
		return true
	}
	argument := vba242QuotedArgument(arguments)
	if argument == "" {
		value := strings.TrimSpace(arguments)
		number, err := strconv.Atoi(value)
		return err == nil && number > 0
	}
	argument = strings.ReplaceAll(argument, "$", "")
	argument = strings.ReplaceAll(argument, " ", "")
	parts := strings.Split(argument, ":")
	if len(parts) > 2 || len(parts) == 0 {
		return false
	}
	for _, part := range parts {
		if part == "" || !vba242Digits(part) {
			return false
		}
		if number, err := strconv.Atoi(part); err != nil || number <= 0 {
			return false
		}
	}
	return true
}

func vba242FullRangeArgument(arguments string) bool {
	argument := vba242QuotedArgument(arguments)
	if argument == "" {
		return false
	}
	argument = strings.ReplaceAll(strings.ToUpper(argument), "$", "")
	argument = strings.ReplaceAll(argument, " ", "")
	if vba242FullColumnArgument(`"`+argument+`"`) || vba242FullRowArgument(`"`+argument+`"`) {
		return true
	}
	return argument == "A1:XFD1048576" || argument == "R1C1:R1048576C16384"
}

func vba242MaxCellsRangeArgument(arguments string) bool {
	parts := splitArgs(arguments)
	if len(parts) != 2 {
		return false
	}
	return vba242CellsCoordinate(parts[0]) == [2]int{1, 1} && vba242CellsCoordinate(parts[1]) == [2]int{1048576, 16384}
}

func vba242CellsCoordinate(expression string) [2]int {
	expression = strings.TrimSpace(expression)
	calls := vba242ScanCalls(expression, map[string]bool{"cells": true})
	if len(calls) != 1 || strings.TrimSpace(expression[:calls[0].start]) != "" || calls[0].close != len(expression)-1 {
		return [2]int{}
	}
	parts := splitArgs(calls[0].arguments)
	if len(parts) != 2 {
		return [2]int{}
	}
	row, rowErr := strconv.Atoi(strings.TrimSpace(parts[0]))
	column, columnErr := strconv.Atoi(strings.TrimSpace(parts[1]))
	if rowErr != nil || columnErr != nil {
		return [2]int{}
	}
	return [2]int{row, column}
}

func vba242RangeKind(arguments string) string {
	argument := strings.ReplaceAll(strings.ToUpper(vba242QuotedArgument(arguments)), "$", "")
	argument = strings.ReplaceAll(argument, " ", "")
	if vba242FullColumnArgument(`"` + argument + `"`) {
		return "column"
	}
	if vba242FullRowArgument(`"` + argument + `"`) {
		return "row"
	}
	return "sheet"
}

func vba242QuotedArgument(arguments string) string {
	value := strings.TrimSpace(arguments)
	if len(value) < 2 || value[0] != '"' || value[len(value)-1] != '"' {
		return ""
	}
	return strings.ReplaceAll(value[1:len(value)-1], `""`, `"`)
}

func vba242ColumnName(value string) bool {
	if len(value) == 0 || len(value) > 3 {
		return false
	}
	for _, r := range value {
		if r < 'A' || r > 'Z' {
			return false
		}
	}
	return true
}

func vba242Digits(value string) bool {
	for _, r := range value {
		if r < '0' || r > '9' {
			return false
		}
	}
	return true
}

func vba242HasExpensiveSink(code string, shapes []vba242Shape) bool {
	if len(shapes) == 0 {
		return false
	}
	calls := vba242ScanCalls(code, vba242SinkMethods)
	for _, shape := range shapes {
		if vba242ShapeBounded(code, shape) {
			continue
		}
		for _, call := range calls {
			if vba242SinkCallAssociated(code, shape, call) {
				return true
			}
		}
	}
	for _, token := range vba242Identifiers(code) {
		if vba242SinkMethods[strings.ToLower(token.name)] && token.previousNonSpace == '.' && token.nextNonSpace != '(' {
			for _, shape := range shapes {
				if vba242ShapeBounded(code, shape) {
					continue
				}
				if vba242SinkTokenAssociated(code, shape, token) {
					return true
				}
			}
		}
		if !vba242SinkProperties[strings.ToLower(token.name)] || token.previousNonSpace != '.' {
			continue
		}
		if strings.TrimSpace(code[token.end:]) == "" {
			continue
		}
		rest := strings.TrimSpace(code[token.end:])
		if strings.HasPrefix(rest, "=") {
			for _, shape := range shapes {
				if shape.start < token.start && vba242SameColonSegment(code, shape.start, token.start) && vba242MemberChainBetween(code, shape.end, token.start) {
					return true
				}
			}
		}
	}
	return false
}

func vba242SinkCallAssociated(code string, shape vba242Shape, call vba242Call) bool {
	if !vba242SameColonSegment(code, shape.start, call.start) {
		return false
	}
	if call.start >= shape.end {
		return vba242MemberChainBetween(code, shape.end, call.start)
	}
	// SetRange accepts the target range as an argument. Other methods such as
	// Find may receive a full range as a search value, which is not the target
	// operation and should not produce VBA242.
	return strings.EqualFold(call.name, "SetRange") && shape.start > call.open && shape.end <= call.close
}

func vba242SinkTokenAssociated(code string, shape vba242Shape, token vba242Identifier) bool {
	if !vba242SameColonSegment(code, shape.start, token.start) {
		return false
	}
	if token.start >= shape.end {
		return vba242MemberChainBetween(code, shape.end, token.start)
	}
	return strings.EqualFold(token.name, "SetRange") && shape.start > token.end
}

func vba242MemberChainBetween(code string, start, end int) bool {
	if start < 0 || end < start || end > len(code) {
		return false
	}
	return vba242MemberChainRe.MatchString(code[start:end])
}

func vba242SameColonSegment(code string, start, end int) bool {
	if start > end {
		start, end = end, start
	}
	depth := 0
	for i := start; i < end && i < len(code); i++ {
		if code[i] == '"' {
			i = vba242SkipString(code, i) - 1
			continue
		}
		switch code[i] {
		case '(':
			depth++
		case ')':
			if depth > 0 {
				depth--
			}
		case ':':
			if depth == 0 {
				return false
			}
		}
	}
	return true
}

func vba242EffectiveShapes(code string, shapes []vba242Shape) []vba242Shape {
	out := make([]vba242Shape, 0, len(shapes))
	for _, shape := range shapes {
		if vba242ShapeBounded(code, shape) {
			continue
		}
		out = append(out, shape)
	}
	return out
}

func vba242ShapeBounded(code string, shape vba242Shape) bool {
	bounders := vba242ScanCalls(code, map[string]bool{"resize": true, "intersect": true})
	for _, bounder := range bounders {
		name := strings.ToLower(bounder.name)
		if name == "intersect" && shape.start > bounder.open && shape.end < bounder.close {
			return true
		}
		if name == "resize" && bounder.start >= shape.end {
			between := strings.TrimSpace(code[shape.end:bounder.start])
			if between == "." || between == "" {
				return true
			}
		}
	}
	return false
}

// vba242Code removes apostrophe/Rem comments but preserves quoted addresses.
func vba242Code(text string) string {
	text = strings.TrimSpace(text)
	bytes := []byte(text)
	inString := false
	for i := 0; i < len(bytes); i++ {
		if bytes[i] == '"' {
			if inString && i+1 < len(bytes) && bytes[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
			continue
		}
		if !inString && i+3 <= len(bytes) && strings.EqualFold(string(bytes[i:i+3]), "rem") && (i == 0 || vba242RemFollowsColon(bytes, i)) && (i+3 == len(bytes) || unicode.IsSpace(rune(bytes[i+3]))) {
			return string(bytes[:i])
		}
		if bytes[i] == '\'' && !inString {
			return string(bytes[:i])
		}
	}
	return string(bytes)
}

func vba242RemFollowsColon(code []byte, index int) bool {
	for i := index - 1; i >= 0; i-- {
		switch code[i] {
		case ' ', '\t', '\r', '\n':
			continue
		default:
			return code[i] == ':'
		}
	}
	return false
}

type vba242Identifier struct {
	name             string
	start            int
	end              int
	previousNonSpace byte
	nextNonSpace     byte
}

func vba242Identifiers(code string) []vba242Identifier {
	var out []vba242Identifier
	for i := 0; i < len(code); {
		if code[i] == '"' {
			i = vba242SkipString(code, i)
			continue
		}
		if !isVBA242IdentifierStart(code[i]) {
			i++
			continue
		}
		start := i
		i++
		for i < len(code) && isVBA242IdentifierPart(code[i]) {
			i++
		}
		end := i
		prev := byte(0)
		for j := start - 1; j >= 0; j-- {
			if code[j] != ' ' && code[j] != '\t' && code[j] != '\r' && code[j] != '\n' {
				prev = code[j]
				break
			}
		}
		next := byte(0)
		for j := end; j < len(code); j++ {
			if code[j] != ' ' && code[j] != '\t' && code[j] != '\r' && code[j] != '\n' {
				next = code[j]
				break
			}
		}
		out = append(out, vba242Identifier{name: code[start:end], start: start, end: end, previousNonSpace: prev, nextNonSpace: next})
	}
	return out
}

func vba242ScanCalls(code string, wanted map[string]bool) []vba242Call {
	var out []vba242Call
	for i := 0; i < len(code); {
		if code[i] == '"' {
			i = vba242SkipString(code, i)
			continue
		}
		if !isVBA242IdentifierStart(code[i]) {
			i++
			continue
		}
		start := i
		i++
		for i < len(code) && isVBA242IdentifierPart(code[i]) {
			i++
		}
		name := code[start:i]
		if !wanted[strings.ToLower(name)] {
			continue
		}
		open := i
		for open < len(code) && (code[open] == ' ' || code[open] == '\t' || code[open] == '\r' || code[open] == '\n') {
			open++
		}
		if open >= len(code) || code[open] != '(' {
			continue
		}
		close := vba242MatchingParen(code, open)
		if close < 0 {
			continue
		}
		out = append(out, vba242Call{name: name, start: start, open: open, close: close, arguments: code[open+1 : close]})
		i = close + 1
	}
	return out
}

func vba242MatchingParen(code string, open int) int {
	depth := 0
	for i := open; i < len(code); i++ {
		if code[i] == '"' {
			i = vba242SkipString(code, i) - 1
			continue
		}
		switch code[i] {
		case '(':
			depth++
		case ')':
			depth--
			if depth == 0 {
				return i
			}
		}
	}
	return -1
}

func vba242SkipString(code string, start int) int {
	for i := start + 1; i < len(code); i++ {
		if code[i] != '"' {
			continue
		}
		if i+1 < len(code) && code[i+1] == '"' {
			i++
			continue
		}
		return i + 1
	}
	return len(code)
}

func isVBA242IdentifierStart(value byte) bool {
	return value == '_' || value >= 'A' && value <= 'Z' || value >= 'a' && value <= 'z'
}

func isVBA242IdentifierPart(value byte) bool {
	return isVBA242IdentifierStart(value) || value >= '0' && value <= '9'
}

func vba242Receiver(code string, memberStart int) string {
	prefix := strings.TrimSpace(code[:memberStart])
	if strings.HasSuffix(prefix, ".") {
		prefix = strings.TrimSpace(strings.TrimSuffix(prefix, "."))
		if prefix == "" {
			// A leading-dot member is resolved through an enclosing With
			// receiver that is not available to this syntax-only scanner.
			return "."
		}
	}
	match := vba242ReceiverRe.FindStringSubmatch(prefix)
	if len(match) == 2 {
		return strings.TrimSpace(match[1])
	}
	return ""
}
