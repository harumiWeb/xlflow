package intel

import (
	"context"
	"fmt"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vbadb"
)

// DefaultMemberDiagnosticContext is the protocol-neutral classification
// attached to default-member diagnostics. Analyzer adapters project the same
// values into the public finding envelope.
type DefaultMemberDiagnosticContext struct {
	Kind            string
	Binding         string
	ExpectedContext string
	Member          string
	Depth           int
}

// DefaultMemberDiagnosticsContext reports implicit value coercion, indexed
// default-member calls, unbound default-member access, and bang notation.
// It deliberately uses the existing revision-scoped document index so batch
// and realtime analysis share declaration shadowing and assignment inference.
func (a Analyzer) DefaultMemberDiagnosticsContext(ctx context.Context, doc Document) ([]Diagnostic, error) {
	if ctx == nil {
		ctx = context.Background()
	}
	if a.DB == nil {
		return nil, nil
	}
	if !a.Config.Analyze.DetectImplicitDefaultMemberAccess &&
		!a.Config.Analyze.DetectUnboundDefaultMemberAccess &&
		!a.Config.Analyze.DetectBangNotation {
		// Default-member classification is an opt-in analysis. Keep its
		// deterministic VBA249 extensions behind the same gate so the default
		// analyzer path does not pay for a source scan or surface a new policy.
		return nil, nil
	}
	resolver, ok := a.NewDocumentExpressionTypeResolver(doc)
	if !ok {
		return nil, nil
	}

	var out []Diagnostic
	seen := make(map[string]struct{})
	appendDiagnostic := func(diagnostic Diagnostic) {
		key := fmt.Sprintf("%s:%d:%d:%d:%d", diagnostic.Code, diagnostic.Range.Start.Line, diagnostic.Range.Start.Character, diagnostic.Range.End.Line, diagnostic.Range.End.Character)
		if _, exists := seen[key]; exists {
			return
		}
		seen[key] = struct{}{}
		out = append(out, diagnostic)
	}

	for index, logicalLine := range logicalLinesForCallAnalysis(doc.Source) {
		if index&0x3f == 0 {
			if err := ctx.Err(); err != nil {
				return nil, err
			}
		}
		if a.Config.Analyze.DetectBangNotation {
			for _, access := range bangMemberAccesses(logicalLine.Text) {
				startLine, startColumn := logicalLine.positionForOffset(access.start)
				endLine, endColumn := logicalLine.positionForOffset(access.end)
				appendDiagnostic(Diagnostic{
					Code: "VBA255", Severity: "information", Source: "xlflow", Rule: "VBA255", Confidence: "high",
					Message:       fmt.Sprintf("Bang notation accesses %q through a stringly typed default member.", access.member),
					Range:         Range{Start: Position{Line: startLine, Character: startColumn}, End: Position{Line: endLine, Character: endColumn}},
					DefaultMember: &DefaultMemberDiagnosticContext{Kind: "bang", Binding: "unbound", ExpectedContext: "value", Member: access.member, Depth: 1},
				})
			}
		}
		if defaultMemberDeclarationLine(logicalLine.Text) {
			continue
		}

		for _, call := range callsOnLine(logicalLine.Text) {
			if !call.Parenthesized || strings.ContainsAny(call.Target, ".!") {
				continue
			}
			target := strings.TrimSpace(call.Target)
			if target == "" {
				continue
			}
			callRange := logicalLine.callRange(call)
			offset := byteOffsetForDocumentPosition(doc, callRange.Start)
			if a.visibleCallableSymbolAtContext(doc, target, offset, resolver.typeContext) {
				continue
			}
			if _, declared := a.visibleSymbolTypeInfoAtContext(doc, target, offset, resolver.typeContext); !declared {
				if _, global := a.DB.ResolveGlobal(target); global {
					// Range(...), Cells(...), and similar host globals are explicit
					// property calls, not calls to a value's default member.
					continue
				}
			}
			receiverType, resolved := resolver.ResolveAt(target, callRange.Start.Line)
			if !resolved {
				continue
			}
			initialized := a.defaultMemberReceiverInitialized(doc, target, offset, resolver.typeContext)
			if diagnostic, ok := a.defaultMemberDiagnostic(receiverType, "indexed", "value", callRange, target, len(call.Arguments), initialized); ok {
				appendDiagnostic(diagnostic)
			}
		}

		assignment := assignmentOperatorIndex(logicalLine.Text)
		if assignment < 0 {
			continue
		}
		setUsed, _ := assignmentLHSExpression(logicalLine.Text[:assignment])
		if setUsed {
			continue
		}
		rhs := strings.TrimSpace(logicalLine.Text[assignment+1:])
		if !implicitValueCandidate(rhs) {
			continue
		}
		rhsStart := strings.Index(logicalLine.Text[assignment+1:], rhs)
		if rhsStart < 0 {
			continue
		}
		rhsStart += assignment + 1
		startLine, startColumn := logicalLine.positionForOffset(rhsStart)
		endLine, endColumn := logicalLine.positionForOffset(rhsStart + len(rhs))
		typ, resolved := resolver.ResolveAt(rhs, startLine)
		if !resolved {
			continue
		}
		offset := byteOffsetForDocumentPosition(doc, Position{Line: startLine, Character: startColumn})
		initialized := a.defaultMemberReceiverInitialized(doc, rhs, offset, resolver.typeContext)
		if diagnostic, ok := a.defaultMemberDiagnostic(typ, "implicit", "value", Range{
			Start: Position{Line: startLine, Character: startColumn},
			End:   Position{Line: endLine, Character: endColumn},
		}, rhs, 0, initialized); ok {
			appendDiagnostic(diagnostic)
		}
	}
	return out, ctx.Err()
}

func (a Analyzer) visibleCallableSymbolAtContext(doc Document, name string, offset int, ctx *documentTypeContext) bool {
	position := positionForDocumentByteOffset(doc, offset)
	procedure := currentProcedureNameAt(doc, position, ctx)
	var symbols []Symbol
	if ctx != nil {
		if ctx.index != nil {
			symbols = ctx.index.symbolsByName[indexName(name)]
		} else {
			symbols = ctx.symbols
		}
	} else if index, ok := a.documentIndexFor(doc); ok {
		symbols = index.symbolsByName[indexName(name)]
	}
	for _, symbol := range symbols {
		if strings.EqualFold(symbol.Name, name) && strings.EqualFold(symbol.Kind, "function") && a.visibleDefinitionSymbol(doc, procedure, symbol) {
			return true
		}
	}
	return false
}

func defaultMemberDeclarationLine(line string) bool {
	line = strings.ToLower(strings.TrimSpace(line))
	for _, prefix := range []string{"dim ", "static ", "const ", "redim ", "public ", "private ", "friend ", "property ", "sub ", "function "} {
		if strings.HasPrefix(line, prefix) {
			return true
		}
	}
	return false
}

func (a Analyzer) defaultMemberDiagnostic(receiverType, kind, expected string, diagnosticRange Range, expression string, argumentCount int, receiverInitialized bool) (Diagnostic, bool) {
	receiverType = canonicalDiagnosticType(a.DB, receiverType)
	if lowConfidenceDiagnosticType(receiverType) {
		if !a.Config.Analyze.DetectUnboundDefaultMemberAccess {
			return Diagnostic{}, false
		}
		return Diagnostic{
			Code: "VBA254", Severity: "information", Source: "xlflow", Rule: "VBA254", Confidence: "medium",
			Message:       fmt.Sprintf("%s may invoke an unbound default member because its runtime object type is not known.", expression),
			Range:         diagnosticRange,
			DefaultMember: &DefaultMemberDiagnosticContext{Kind: kind, Binding: "unbound", ExpectedContext: expected, Depth: 1},
		}, true
	}

	member, depth, resultType, status := a.resolveDefaultMemberValueChain(receiverType, argumentCount)
	switch status {
	case defaultMemberKnown:
		if !a.Config.Analyze.DetectImplicitDefaultMemberAccess {
			return Diagnostic{}, false
		}
		if depth > 1 {
			kind = "recursive"
		}
		return Diagnostic{
			Code: "VBA253", Severity: "warning", Source: "xlflow", Rule: "VBA253", Confidence: "high",
			Message:       fmt.Sprintf("%s implicitly invokes default member %q and produces %s.", expression, member, resultType),
			Range:         diagnosticRange,
			DefaultMember: &DefaultMemberDiagnosticContext{Kind: kind, Binding: "known", ExpectedContext: expected, Member: member, Depth: depth},
		}, true
	case defaultMemberUnbound:
		if !a.Config.Analyze.DetectUnboundDefaultMemberAccess {
			return Diagnostic{}, false
		}
		return Diagnostic{
			Code: "VBA254", Severity: "information", Source: "xlflow", Rule: "VBA254", Confidence: "medium",
			Message:       fmt.Sprintf("%s may require an unbound default-member chain to produce a value.", expression),
			Range:         diagnosticRange,
			DefaultMember: &DefaultMemberDiagnosticContext{Kind: kind, Binding: "unbound", ExpectedContext: expected, Member: member, Depth: max(1, depth)},
		}, true
	case defaultMemberInvalid:
		if !a.Config.Analyze.DetectDeterministicRuntimeErrors || !receiverInitialized {
			return Diagnostic{}, false
		}
		return Diagnostic{
			Code: "VBA249", Severity: "error", Source: "xlflow", Rule: "VBA249", Confidence: "high",
			Message:       fmt.Sprintf("%s cannot produce the required value because %s has no usable default member.", expression, receiverType),
			Range:         diagnosticRange,
			DefaultMember: &DefaultMemberDiagnosticContext{Kind: kind, Binding: "invalid", ExpectedContext: expected, Member: member, Depth: max(1, depth)},
		}, true
	case defaultMemberCycle:
		if !a.Config.Analyze.DetectDeterministicRuntimeErrors || !receiverInitialized {
			return Diagnostic{}, false
		}
		return Diagnostic{
			Code: "VBA249", Severity: "error", Source: "xlflow", Rule: "VBA249", Confidence: "high",
			Message:       fmt.Sprintf("%s cannot produce the required value because its default-member chain is cyclic.", expression),
			Range:         diagnosticRange,
			DefaultMember: &DefaultMemberDiagnosticContext{Kind: "recursive", Binding: "invalid", ExpectedContext: expected, Member: member, Depth: max(1, depth)},
		}, true
	default:
		return Diagnostic{}, false
	}
}

func (a Analyzer) defaultMemberReceiverInitialized(doc Document, name string, offset int, ctx *documentTypeContext) bool {
	index := (*documentIndex)(nil)
	if ctx != nil {
		index = ctx.index
	}
	if index == nil {
		var ok bool
		index, ok = a.documentIndexFor(doc)
		if !ok || index == nil {
			return false
		}
	}
	position := positionForDocumentByteOffset(doc, offset)
	procedure := currentProcedureNameAt(doc, position, ctx)
	assignment, ok := index.nearestAssignment(name, procedure, position)
	if !ok {
		return false
	}
	if match := newAssignmentExprRe.FindStringSubmatch(assignment.expression); len(match) == 2 {
		return true
	}
	return len(createObjectExprRe.FindStringSubmatch(assignment.expression)) == 2
}

type defaultMemberStatus uint8

const (
	defaultMemberUnknown defaultMemberStatus = iota
	defaultMemberKnown
	defaultMemberUnbound
	defaultMemberInvalid
	defaultMemberCycle
)

func (a Analyzer) resolveDefaultMemberValueChain(typeName string, initialArgumentCount int) (member string, depth int, resultType string, status defaultMemberStatus) {
	seen := make(map[string]struct{})
	current := canonicalDiagnosticType(a.DB, typeName)
	for current != "" {
		if lowConfidenceDiagnosticType(current) {
			if depth > 0 {
				return member, depth, current, defaultMemberKnown
			}
			return member, depth, current, defaultMemberUnbound
		}
		if valueDefaultMemberType(current) {
			if depth == 0 {
				return "", 0, current, defaultMemberUnknown
			}
			return member, depth, current, defaultMemberKnown
		}
		key := strings.ToLower(current)
		if _, exists := seen[key]; exists {
			return member, depth, current, defaultMemberCycle
		}
		seen[key] = struct{}{}
		if a.ProjectDefaultTypes[key] {
			candidate, ok := a.ProjectDefaultMembers[key]
			if !ok {
				return member, depth, current, defaultMemberInvalid
			}
			argumentCount := 0
			if depth == 0 {
				argumentCount = initialArgumentCount
			}
			if !defaultMemberAcceptsArguments(candidate, argumentCount) {
				return member, depth, current, defaultMemberUnknown
			}
			if depth == 0 {
				member = candidate.Name
			}
			depth++
			if candidate.ReturnType == "" {
				return member, depth, "", defaultMemberUnbound
			}
			current = canonicalDiagnosticType(a.DB, candidate.ReturnType)
			continue
		}
		typ, ok := a.DB.ResolveType(current)
		if !ok {
			return member, depth, current, defaultMemberUnbound
		}
		if valueDefaultMemberKind(typ.Kind) {
			if depth == 0 {
				return "", 0, current, defaultMemberUnknown
			}
			return member, depth, current, defaultMemberKnown
		}
		candidate, ok := uniqueDefaultMember(typ)
		if !ok {
			if strings.EqualFold(typ.Source, "typelib") && strings.EqualFold(typ.Confidence, "generated") && !a.TypeDBResolutionIncomplete {
				return member, depth, current, defaultMemberInvalid
			}
			return member, depth, current, defaultMemberUnknown
		}
		argumentCount := 0
		if depth == 0 {
			argumentCount = initialArgumentCount
		}
		if !defaultMemberAcceptsArguments(candidate, argumentCount) {
			return member, depth, current, defaultMemberUnknown
		}
		if depth == 0 {
			member = candidate.Name
		}
		depth++
		if candidate.ReturnType == "" {
			return member, depth, "", defaultMemberUnbound
		}
		current = canonicalDiagnosticType(a.DB, candidate.ReturnType)
	}
	return member, depth, current, defaultMemberUnbound
}

func defaultMemberAcceptsArguments(member vbadb.MemberInfo, argumentCount int) bool {
	minimum := 0
	maximum := len(member.Parameters)
	for _, parameter := range member.Parameters {
		if parameter.ParamArray {
			maximum = -1
			continue
		}
		if !parameter.Optional {
			minimum++
		}
	}
	return argumentCount >= minimum && (maximum < 0 || argumentCount <= maximum)
}

func uniqueDefaultMember(typ vbadb.TypeInfo) (vbadb.MemberInfo, bool) {
	var candidates []vbadb.MemberInfo
	for _, candidate := range append(append([]vbadb.MemberInfo{}, typ.Properties...), typ.Methods...) {
		if candidate.Default || typ.DefaultMember != "" && strings.EqualFold(candidate.Name, typ.DefaultMember) {
			candidates = append(candidates, candidate)
		}
	}
	if len(candidates) == 0 && typ.DefaultMember != "" {
		return vbadb.MemberInfo{Name: typ.DefaultMember, ReturnType: typ.DefaultMemberType, Default: true}, true
	}
	if len(candidates) != 1 {
		return vbadb.MemberInfo{}, false
	}
	return candidates[0], true
}

func valueDefaultMemberType(typeName string) bool {
	switch strings.ToLower(strings.TrimSpace(typeName)) {
	case "boolean", "byte", "currency", "date", "decimal", "double", "integer", "long", "longlong", "longptr", "single", "string":
		return true
	default:
		return false
	}
}

func valueDefaultMemberKind(kind string) bool {
	return strings.EqualFold(strings.TrimSpace(kind), "enum")
}

func implicitValueCandidate(expression string) bool {
	expression = strings.TrimSpace(expression)
	if expression == "" || strings.ContainsAny(expression, "!()") {
		return false
	}
	for _, r := range expression {
		if r == '.' || r == '_' || r >= '0' && r <= '9' || r >= 'A' && r <= 'Z' || r >= 'a' && r <= 'z' {
			continue
		}
		return false
	}
	// An explicit member already names the value/property being requested.
	return !strings.Contains(expression, ".")
}

type bangMemberAccess struct {
	member     string
	start, end int
}

func bangMemberAccesses(line string) []bangMemberAccess {
	if commentStart := commentStartIndex(line); commentStart >= 0 {
		line = line[:commentStart]
	}
	masked := codeWithoutStringLiterals(line)
	trimmed := strings.TrimSpace(masked)
	if strings.EqualFold(trimmed, "rem") || len(trimmed) > 4 && strings.EqualFold(trimmed[:4], "rem ") {
		return nil
	}
	var out []bangMemberAccess
	for index := 1; index+1 < len(masked); index++ {
		if masked[index] != '!' {
			continue
		}
		left := index - 1
		for left >= 0 && (masked[left] == ' ' || masked[left] == '\t') {
			left--
		}
		right := index + 1
		for right < len(masked) && (masked[right] == ' ' || masked[right] == '\t') {
			right++
		}
		if left < 0 || right >= len(masked) || (!isIdentRune(rune(masked[left])) && masked[left] != ')') || !isIdentStartRune(rune(masked[right])) {
			continue
		}
		end := right + 1
		for end < len(masked) && isIdentRune(rune(masked[end])) {
			end++
		}
		out = append(out, bangMemberAccess{member: line[right:end], start: index, end: end})
	}
	return out
}
