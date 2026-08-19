package intel

import (
	"context"
	"fmt"
	"regexp"
	"strings"

	staticrules "github.com/harumiWeb/xlflow/internal/staticanalysis/rules"
	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
)

var (
	memberExpressionPattern = regexp.MustCompile(`[A-Za-z0-9_)\]]\.[A-Za-z_]`)
	ptrSafeDeclarePattern   = regexp.MustCompile(`(?i)\bdeclare\s+ptrsafe\b`)
	numericByRefLiteral     = regexp.MustCompile(`^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eEdD][+-]?\d+)?[&^!#@]?$`)
)

// ByRefArgumentDiagnostics reports unsafe calls to resolved project-local
// procedures. It deliberately ignores unresolved, ambiguous, external, and
// late-bound calls: their signatures are not evidence strong enough to make a
// ByRef compatibility claim.
func (a Analyzer) ByRefArgumentDiagnostics(doc Document) []Diagnostic {
	return a.ByRefArgumentDiagnosticsContext(context.Background(), doc)
}

// ByRefArgumentDiagnosticsContext is the cancellable form used by realtime LSP analysis.
func (a Analyzer) ByRefArgumentDiagnosticsContext(ctx context.Context, doc Document) []Diagnostic {
	if recorder := analysisstats.FromContext(ctx); recorder != nil {
		recorder.Add("byref_diagnostic_passes", 1)
	}
	// Resolve calls in the current module directly from its immutable snapshot.
	// A newly opened document's workspace overlay is intentionally absent while
	// background analysis is pending, but file-local diagnostics must still be
	// complete. Build this list once rather than cloning it for every call site.
	localSymbols, err := a.DocumentSymbolsContext(ctx, doc)
	if err != nil || ctx.Err() != nil {
		return nil
	}
	localSymbolsByName := make(map[string][]Symbol, len(localSymbols))
	for _, symbol := range localSymbols {
		key := strings.ToLower(strings.TrimSpace(symbol.Name))
		localSymbolsByName[key] = append(localSymbolsByName[key], symbol)
	}
	var out []Diagnostic
	for i, logicalLine := range logicalLinesForCallAnalysis(doc.Source) {
		if i&0x3f == 0 && ctx.Err() != nil {
			return nil
		}
		calls := callsOnLine(logicalLine.Text)
		for _, call := range calls {
			if byRefCallIsShadowedByWholeArgumentForm(call, calls) {
				continue
			}
			callRange := logicalLine.callRange(call)
			call.DiagnosticRange = &callRange
			sig, resolved, err := a.resolveProjectLocalCallSignature(doc, localSymbolsByName, call.Target, callRange.Start)
			if err != nil || !resolved || !sig.projectLocal {
				continue
			}
			positional := 0
			for _, arg := range call.Arguments {
				param, next, ok := signatureParameterForArgument(sig.Parameters, arg, positional)
				if arg.Name == "" {
					positional = next
				}
				if !ok || param.ParamArray || !isByRefParameter(param) {
					continue
				}
				if diagnostic, found := a.byRefArgumentDiagnostic(doc, callRange.Start, callRange.Start.Line, call, arg.Text, param, sig.declaringModule); found && (diagnostic.Code != "VBA206" || a.Config.Analyze.DetectByRefArgumentMismatch) {
					out = append(out, diagnostic)
				}
			}
		}
	}
	if ctx.Err() != nil {
		return nil
	}
	if a.Config.Analyze.DetectByRefArgumentMismatch {
		out = append(out, a.ptrSafeDeclareDiagnostics(doc)...)
	}
	return out
}

// callsOnLine retains both parenthesized and parenthesis-free VBA call forms
// for editor features. For `TakeLong (value)`, that produces two equivalent
// calls: one with `value` and one with `(value)`. VBA206 must prefer the
// latter, because only it preserves the important temporary-value semantics.
func byRefCallIsShadowedByWholeArgumentForm(call parsedCall, calls []parsedCall) bool {
	if len(call.Arguments) != 1 || hasWholeExpressionParentheses(strings.TrimSpace(call.Arguments[0].Text)) {
		return false
	}
	for _, other := range calls {
		if other.End != call.End || !strings.EqualFold(other.Target, call.Target) || len(other.Arguments) != 1 {
			continue
		}
		wrapped := strings.TrimSpace(other.Arguments[0].Text)
		if hasWholeExpressionParentheses(wrapped) && strings.EqualFold(strings.TrimSpace(wrapped[1:len(wrapped)-1]), strings.TrimSpace(call.Arguments[0].Text)) {
			return true
		}
	}
	return false
}

// resolveProjectLocalCallSignature deliberately resolves only one project
// procedure symbol. Unlike general signature help, it neither falls back to a
// built-in/member signature nor selects an arbitrary overload: VBA206 needs a
// concrete callee declaration before it can make a ByRef claim.
func (a Analyzer) resolveProjectLocalCallSignature(doc Document, localSymbolsByName map[string][]Symbol, target string, pos Position) (Signature, bool, error) {
	target = strings.TrimSpace(target)
	if target == "" || strings.HasPrefix(target, ".") {
		return Signature{}, false, nil
	}
	receiver, member, qualified := splitCallTarget(target)
	query := target
	if qualified {
		query = member
	}
	currentProcedure := currentProcedureNameForDocument(doc, pos)
	module := moduleNameForDocument(doc)
	localCandidates := localSymbolsByName[strings.ToLower(strings.TrimSpace(query))]
	if !qualified && nonCallableLocalShadowsProjectCall(a, doc, currentProcedure, localCandidates, target) {
		return Signature{}, false, nil
	}
	localMatches := matchingProjectCallSymbols(a, doc, currentProcedure, localCandidates, target, receiver, member, qualified)
	if !qualified {
		localMatches = symbolsInModule(localMatches, module)
	}
	if len(localMatches) == 1 {
		return signatureFromSymbol(localMatches[0]), true, nil
	}
	if len(localMatches) > 1 {
		return Signature{}, false, nil
	}
	syms, err := a.WorkspaceSymbolsQuery([]Document{doc}, WorkspaceSymbolQuery{Text: query, Mode: WorkspaceSymbolQueryExact})
	if err != nil {
		return Signature{}, false, err
	}
	matches := matchingProjectCallSymbols(a, doc, currentProcedure, syms, target, receiver, member, qualified)
	if !qualified {
		local := symbolsInModule(matches, module)
		if len(local) == 1 {
			return signatureFromSymbol(local[0]), true, nil
		}
		if len(local) > 1 {
			return Signature{}, false, nil
		}
	}
	if len(matches) != 1 {
		return Signature{}, false, nil
	}
	return signatureFromSymbol(matches[0]), true, nil
}

func nonCallableLocalShadowsProjectCall(a Analyzer, doc Document, currentProcedure string, syms []Symbol, target string) bool {
	for _, sym := range syms {
		if !strings.EqualFold(sym.Name, target) || !symbolCanShadowProjectCall(sym) {
			continue
		}
		if a.visibleCompletionSymbol(doc, currentProcedure, sym) {
			return true
		}
	}
	return false
}

func symbolCanShadowProjectCall(sym Symbol) bool {
	switch strings.ToLower(strings.TrimSpace(sym.Kind)) {
	case "const", "field", "local_variable", "module_variable", "parameter", "withevents_field":
		return true
	default:
		return false
	}
}

func matchingProjectCallSymbols(a Analyzer, doc Document, currentProcedure string, syms []Symbol, target, receiver, member string, qualified bool) []Symbol {
	matches := make([]Symbol, 0, len(syms))
	for _, sym := range syms {
		if !callableCompletionSymbol(sym) || !a.visibleCompletionSymbol(doc, currentProcedure, sym) {
			continue
		}
		if qualified {
			if !strings.EqualFold(sym.Name, member) || !strings.EqualFold(sym.Module, receiver) {
				continue
			}
		} else if !strings.EqualFold(sym.Name, target) {
			continue
		}
		matches = append(matches, sym)
	}
	return matches
}

func symbolsInModule(syms []Symbol, module string) []Symbol {
	local := make([]Symbol, 0, len(syms))
	for _, sym := range syms {
		if strings.EqualFold(sym.Module, module) {
			local = append(local, sym)
		}
	}
	return local
}

func isByRefParameter(param Parameter) bool {
	return !strings.EqualFold(strings.TrimSpace(param.Passing), "ByVal")
}

func (a Analyzer) byRefArgumentDiagnostic(doc Document, pos Position, lineNo int, call parsedCall, text string, param Parameter, declaringModule string) (Diagnostic, bool) {
	expr := strings.TrimSpace(text)
	if expr == "" {
		return Diagnostic{}, false
	}
	if hasWholeExpressionParentheses(expr) {
		return byRefDiagnostic(lineNo, call, fmt.Sprintf("Argument `%s` for ByRef parameter `%s` is parenthesized. VBA evaluates it into a temporary value, so changes made by the procedure do not update the original argument.", expr, param.Name)), true
	}
	if _, literal := byRefLiteralType(expr); literal {
		return byRefDiagnostic(lineNo, call, fmt.Sprintf("Argument `%s` for ByRef parameter `%s` is a literal. Pass a writable variable because procedure changes cannot update a literal.", expr, param.Name)), true
	}
	if strings.HasPrefix(strings.ToLower(expr), "new ") {
		return byRefDiagnostic(lineNo, call, fmt.Sprintf("Argument `%s` for ByRef parameter `%s` is a newly created object expression. Pass a writable variable to observe any ByRef replacement.", expr, param.Name)), true
	}
	if isIdentifier(expr) {
		inferred, ok := a.inferWordTypeInfoAt(doc, expr, byteOffsetForDocumentPosition(doc, pos))
		if !ok || lowConfidenceDiagnosticType(inferred.Type) {
			return Diagnostic{}, false
		}
		if byRefTypesMismatch(inferred.Type, inferred.IsArray, param.Type, param.IsArray, declaringModule) {
			return byRefTypeMismatchDiagnostic(lineNo, call, fmt.Sprintf("Argument `%s` has type %s, but ByRef parameter `%s` requires %s.", expr, displayInferredType(inferred), param.Name, displayParameterType(param))), true
		}
		return Diagnostic{}, false
	}
	if looksMemberExpression(expr) {
		return byRefDiagnostic(lineNo, call, fmt.Sprintf("Argument `%s` for ByRef parameter `%s` is a property or member expression. Any mutation is indirect and may be surprising; pass a writable variable instead.", expr, param.Name)), true
	}
	if looksIndexedExpression(expr) {
		return byRefDiagnostic(lineNo, call, fmt.Sprintf("Argument `%s` for ByRef parameter `%s` is an array element or indexed expression. Any mutation is indirect and may be surprising; pass a writable variable instead.", expr, param.Name)), true
	}
	return byRefDiagnostic(lineNo, call, fmt.Sprintf("Argument `%s` for ByRef parameter `%s` is an expression rather than a writable variable. VBA may pass a temporary value, so procedure changes can be lost.", expr, param.Name)), true
}

func byRefDiagnostic(lineNo int, call parsedCall, message string) Diagnostic {
	diagnostic := callDiagnostic(lineNo, call, message)
	diagnostic.Code = "VBA206"
	diagnostic.Rule = "VBA206"
	diagnostic.Confidence = "high"
	return diagnostic
}

func byRefTypeMismatchDiagnostic(lineNo int, call parsedCall, message string) Diagnostic {
	diagnostic := byRefDiagnostic(lineNo, call, message)
	diagnostic.Code = "VBA228"
	diagnostic.Rule = "VBA228"
	if metadata, ok := staticrules.Lookup(diagnostic.Code); ok {
		diagnostic.Severity = string(metadata.DefaultSeverity)
	}
	return diagnostic
}

func hasWholeExpressionParentheses(expr string) bool {
	if len(expr) < 2 || expr[0] != '(' || expr[len(expr)-1] != ')' {
		return false
	}
	return matchingParen(expr, 0) == len(expr)-1
}

func byRefLiteralType(expr string) (string, bool) {
	lower := strings.ToLower(strings.TrimSpace(expr))
	if strings.HasPrefix(expr, `"`) && strings.HasSuffix(expr, `"`) {
		return "String", true
	}
	switch lower {
	case "true", "false":
		return "Boolean", true
	case "nothing":
		return "Object", true
	case "null", "empty":
		return "Variant", true
	}
	if isNumericByRefLiteral(expr) {
		return numericByRefLiteralType(expr), true
	}
	return "", false
}

func isNumericByRefLiteral(expr string) bool {
	return numericByRefLiteral.MatchString(strings.TrimSpace(expr))
}

func numericByRefLiteralType(expr string) string {
	lower := strings.ToLower(strings.TrimSpace(expr))
	switch {
	case strings.Contains(lower, "@"):
		return "Currency"
	case strings.Contains(lower, "#") || strings.Contains(lower, ".") || strings.ContainsAny(lower, "ed"):
		return "Double"
	case strings.Contains(lower, "!"):
		return "Single"
	case strings.Contains(lower, "^"):
		return "LongLong"
	default:
		return "Long"
	}
}

func looksMemberExpression(expr string) bool { return memberExpressionPattern.MatchString(expr) }

func looksIndexedExpression(expr string) bool {
	open := strings.Index(expr, "(")
	return open > 0 && strings.HasSuffix(strings.TrimSpace(expr), ")")
}

func byRefTypesMismatch(actual string, actualArray bool, expected string, expectedArray bool, declaringModule string) bool {
	actual = byRefCanonicalType(actual)
	expected = byRefCanonicalType(expected)
	if actual == "" || expected == "" || actual == "variant" || actual == "object" || actual == "any" || expected == "variant" || expected == "object" || expected == "any" {
		return false
	}
	if actualArray != expectedArray {
		return true
	}
	if byRefQualifiedTypeMatchesDeclaringModule(actual, expected, declaringModule) {
		return false
	}
	return actual != expected
}

func byRefQualifiedTypeMatchesDeclaringModule(actual, expected, declaringModule string) bool {
	if strings.Contains(expected, ".") || strings.TrimSpace(declaringModule) == "" {
		return false
	}
	separator := strings.LastIndex(actual, ".")
	if separator <= 0 || separator == len(actual)-1 {
		return false
	}
	return strings.EqualFold(actual[:separator], strings.TrimSpace(declaringModule)) && actual[separator+1:] == expected
}

func byRefCanonicalType(typ string) string {
	typ = strings.TrimSpace(strings.ToLower(typ))
	typ = strings.TrimPrefix(typ, "vba.")
	typ = strings.TrimPrefix(typ, "excel.")
	return typ
}

type winAPIPointerContract struct {
	returnPointer bool
	parameters    map[string]bool
}

var winAPIPointerContracts = map[string]winAPIPointerContract{
	"findwindow":       {returnPointer: true},
	"findwindowex":     {returnPointer: true, parameters: map[string]bool{"hwndparent": true, "hwndchildafter": true}},
	"getwindowlongptr": {returnPointer: true, parameters: map[string]bool{"hwnd": true}},
	"setwindowlongptr": {returnPointer: true, parameters: map[string]bool{"hwnd": true, "dwnewlong": true}},
	"getclasslongptr":  {returnPointer: true, parameters: map[string]bool{"hwnd": true}},
	"setclasslongptr":  {returnPointer: true, parameters: map[string]bool{"hwnd": true, "dwnewlong": true}},
	"getmodulehandle":  {returnPointer: true},
	"getprocaddress":   {returnPointer: true, parameters: map[string]bool{"hmodule": true}},
	"sendmessage":      {returnPointer: true, parameters: map[string]bool{"hwnd": true, "wparam": true, "lparam": true}},
	"postmessage":      {parameters: map[string]bool{"hwnd": true, "wparam": true, "lparam": true}},
}

func (a Analyzer) ptrSafeDeclareDiagnostics(doc Document) []Diagnostic {
	symbols, err := a.DocumentSymbols(doc)
	if err != nil {
		return nil
	}
	lines := normalizedLines(doc.Source)
	var out []Diagnostic
	for _, symbol := range symbols {
		if !strings.EqualFold(symbol.Kind, "declare_function") && !strings.EqualFold(symbol.Kind, "declare_sub") {
			continue
		}
		line := symbol.Range.Start.Line
		if line < 0 || line >= len(lines) {
			continue
		}
		declaration := lines[line]
		if end := min(len(lines)-1, symbol.Range.End.Line); end > line {
			declaration = strings.Join(lines[line:end+1], " ")
		}
		if !ptrSafeDeclarePattern.MatchString(declaration) {
			continue
		}
		contract := winAPIPointerContracts[strings.ToLower(symbol.Name)]
		if contract.returnPointer && longUsedForPointer(symbol.ReturnType) {
			out = append(out, byRefDeclareDiagnostic(line, lines[line], fmt.Sprintf("PtrSafe Declare `%s` returns a pointer-sized value but declares `Long`. Use `LongPtr`.", symbol.Name), "high"))
		}
		for _, parameter := range symbol.Parameters {
			pointerLike, confidence := pointerLikeDeclareParameter(parameter.Name, contract)
			if !pointerLike || !longUsedForPointer(parameter.Type) {
				continue
			}
			out = append(out, byRefDeclareDiagnostic(line, lines[line], fmt.Sprintf("PtrSafe Declare `%s` parameter `%s` is pointer-sized but declares `Long`. Use `LongPtr`.", symbol.Name, parameter.Name), confidence))
		}
	}
	return out
}

func pointerLikeDeclareParameter(name string, contract winAPIPointerContract) (bool, string) {
	if contract.parameters != nil && contract.parameters[strings.ToLower(name)] {
		return true, "high"
	}
	switch strings.ToLower(strings.TrimSpace(name)) {
	case "hwnd", "hinstance", "hmodule", "hicon", "hcursor", "hmenu", "hbitmap", "hbrush", "handle", "wparam", "lparam", "dwnewlong", "lpaddress":
		return true, "medium"
	default:
		return false, ""
	}
}

func longUsedForPointer(typ string) bool {
	return byRefCanonicalType(typ) == "long"
}

func byRefDeclareDiagnostic(lineNo int, line, message, confidence string) Diagnostic {
	return Diagnostic{
		Code:       "VBA206",
		Severity:   "warning",
		Source:     "xlflow",
		Rule:       "VBA206",
		Confidence: confidence,
		Message:    message,
		Range:      Range{Start: Position{Line: lineNo}, End: Position{Line: lineNo, Character: utf16Len(line)}},
	}
}
