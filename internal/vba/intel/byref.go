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
	localUserDefinedTypes, workspaceUserDefinedTypes, workspaceUserDefinedTypesComplete := a.byRefUserDefinedTypes(doc, localSymbols)
	conditionalLines := conditionalCompilationLines(doc.Source)
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
			sig, resolved, err := a.resolveProjectLocalCallSignature(doc, localSymbolsByName, call.Target, callRange.Start, conditionalLines)
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
				if diagnostic, found := a.byRefArgumentDiagnostic(doc, callRange.Start, callRange.Start.Line, call, arg.Text, param, sig.declaringModule, localUserDefinedTypes, workspaceUserDefinedTypes, workspaceUserDefinedTypesComplete); found && (diagnostic.Code != "VBA206" || a.Config.Analyze.DetectByRefArgumentMismatch) {
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
func (a Analyzer) resolveProjectLocalCallSignature(doc Document, localSymbolsByName map[string][]Symbol, target string, pos Position, conditionalLines map[int]bool) (Signature, bool, error) {
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
	localCandidates := localSymbolsByName[strings.ToLower(strings.TrimSpace(query))]
	if qualified && nonCallableLocalShadowsProjectCall(a, doc, currentProcedure, localSymbolsByName[strings.ToLower(strings.TrimSpace(receiver))], receiver) {
		return Signature{}, false, nil
	}
	if !qualified && nonCallableLocalShadowsProjectCall(a, doc, currentProcedure, localCandidates, target) {
		return Signature{}, false, nil
	}
	localMatches := matchingProjectCallSymbols(a, doc, currentProcedure, localCandidates, target, receiver, member, qualified)
	if !qualified {
		localMatches = a.symbolsInCurrentModule(doc, localMatches)
	}
	if len(localMatches) == 1 {
		if a.conditionallyCompiledCallSymbol(doc, localMatches[0], conditionalLines) {
			return Signature{}, false, nil
		}
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
		local := a.symbolsInCurrentModule(doc, matches)
		if len(local) == 1 {
			if a.conditionallyCompiledCallSymbol(doc, local[0], conditionalLines) {
				return Signature{}, false, nil
			}
			return signatureFromSymbol(local[0]), true, nil
		}
		if len(local) > 1 {
			return Signature{}, false, nil
		}
	}
	if len(matches) != 1 {
		return Signature{}, false, nil
	}
	if a.conditionallyCompiledCallSymbol(doc, matches[0], conditionalLines) {
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

func (a Analyzer) isCurrentModuleSymbol(doc Document, sym Symbol) bool {
	if a.sameDocumentSymbol(doc, sym) {
		return true
	}
	module := moduleNameForCurrentInstance(doc)
	return module != "" && strings.EqualFold(strings.TrimSpace(sym.Module), module)
}

func (a Analyzer) symbolsInCurrentModule(doc Document, syms []Symbol) []Symbol {
	local := make([]Symbol, 0, len(syms))
	for _, sym := range syms {
		if a.isCurrentModuleSymbol(doc, sym) {
			local = append(local, sym)
		}
	}
	return local
}

func (a Analyzer) conditionallyCompiledCurrentModuleSymbol(doc Document, sym Symbol, conditionalLines map[int]bool) bool {
	if len(conditionalLines) == 0 || !a.isCurrentModuleSymbol(doc, sym) {
		return false
	}
	return conditionalLines[sym.Range.Start.Line]
}

func (a Analyzer) conditionallyCompiledCallSymbol(doc Document, sym Symbol, conditionalLines map[int]bool) bool {
	return len(sym.ConditionalBranches) > 0 || a.conditionallyCompiledCurrentModuleSymbol(doc, sym, conditionalLines)
}

func conditionalDirective(line string) string {
	fields := strings.Fields(strings.ToLower(strings.TrimSpace(line)))
	if len(fields) == 0 {
		return ""
	}
	switch fields[0] {
	case "#if", "#elseif", "#endif":
		return fields[0]
	case "#else":
		if len(fields) == 1 {
			return fields[0]
		}
	case "#end":
		if len(fields) == 2 && fields[1] == "if" {
			return "#endif"
		}
	}
	return ""
}

func conditionalCompilationLines(source string) map[int]bool {
	var active map[int]bool
	depth := 0
	for lineNumber, line := range normalizedLines(source) {
		if depth > 0 {
			if active == nil {
				active = make(map[int]bool)
			}
			active[lineNumber] = true
		}
		directive := conditionalDirective(stripLineComment(line))
		switch directive {
		case "#if":
			depth++
		case "#elseif", "#else":
			// The enclosing conditional remains active for every branch.
		case "#endif":
			if depth > 0 {
				depth--
			}
		}
	}
	return active
}

func isByRefParameter(param Parameter) bool {
	return !strings.EqualFold(strings.TrimSpace(param.Passing), "ByVal")
}

func (a Analyzer) byRefArgumentDiagnostic(doc Document, pos Position, lineNo int, call parsedCall, text string, param Parameter, declaringModule string, localUserDefinedTypes map[string]struct{}, workspaceUserDefinedTypes *WorkspaceUserDefinedTypeIndex, workspaceUserDefinedTypesComplete bool) (Diagnostic, bool) {
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
		mismatch := byRefTypesMismatch(inferred.Type, inferred.IsArray, param.Type, param.IsArray, declaringModule)
		if mismatch && !byRefArrayReinterpretationWithWorkspace(inferred, param, localUserDefinedTypes, workspaceUserDefinedTypes) && !a.byRefArrayReinterpretationPending(inferred, param, localUserDefinedTypes, workspaceUserDefinedTypes, workspaceUserDefinedTypesComplete) {
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

func (a Analyzer) byRefUserDefinedTypes(doc Document, localSymbols []Symbol) (map[string]struct{}, *WorkspaceUserDefinedTypeIndex, bool) {
	types := a.byRefLocalUserDefinedTypesForDocument(doc, localSymbols)
	if a.WorkspaceUserDefinedTypes != nil {
		return types, a.WorkspaceUserDefinedTypes, a.WorkspaceUserDefinedTypesComplete
	}
	if a.WorkspaceSymbolQueryFunc == nil && a.WorkspaceSymbolQueryContextFunc == nil && a.WorkspaceSymbolsFunc == nil {
		return types, nil, true
	}
	query := WorkspaceSymbolQuery{
		Text: "type",
		Mode: WorkspaceSymbolQueryKind,
	}
	// WorkspaceSymbolsFunc is the pre-query provider contract and receives only
	// a contains-search string. An empty search is its all-symbols operation;
	// filter those results by kind locally instead of asking it to find the
	// literal word "type".
	if a.WorkspaceSymbolQueryFunc == nil && a.WorkspaceSymbolQueryContextFunc == nil && a.WorkspaceSymbolsFunc != nil {
		query.Text = ""
	}
	projectTypes, err := a.WorkspaceSymbolsQuery([]Document{doc}, query)
	if err != nil {
		return types, nil, false
	}
	return types, NewWorkspaceUserDefinedTypeIndexForProject(projectTypes, a.Config.Project.Name), true
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

// byRefArrayReinterpretation is the VBA/VBE exception for low-level pointer
// access: an array of a project-local user-defined type can be passed to a
// pointer-sized intrinsic array ByRef. The compiler accepts this deliberate
// representation trick even though the declared element types differ.
// Keep the exception structural so it does not depend on a library's type
// names. Non-pointer array mismatches remain VBA228 findings.
func byRefArrayReinterpretation(inferred inferredType, param Parameter, localUserDefinedTypes map[string]struct{}) bool {
	if !inferred.IsArray || !param.IsArray || !isByRefPointerSizedType(param.Type) {
		return false
	}
	// Keep an explicit namespace while checking whether the inferred type is
	// one of the current project's UDTs. byRefCanonicalType intentionally
	// normalizes VBA./Excel. for ordinary compatibility checks, but doing that
	// here could turn an external Excel.Range into a local Range UDT.
	typeName := byRefLocalTypeIdentity(inferred.Type)
	if strings.Contains(typeName, ".") {
		_, ok := localUserDefinedTypes[typeName]
		return ok
	}
	_, ok := localUserDefinedTypes[byRefShortLocalTypeName(typeName)]
	return ok
}

func byRefArrayReinterpretationWithWorkspace(inferred inferredType, param Parameter, localUserDefinedTypes map[string]struct{}, workspaceUserDefinedTypes *WorkspaceUserDefinedTypeIndex) bool {
	if byRefArrayReinterpretation(inferred, param, localUserDefinedTypes) {
		return true
	}
	if !inferred.IsArray || !param.IsArray || !isByRefPointerSizedType(param.Type) || workspaceUserDefinedTypes == nil {
		return false
	}
	return workspaceUserDefinedTypes.matches(inferred.Type)
}

func (a Analyzer) byRefArrayReinterpretationPending(inferred inferredType, param Parameter, localUserDefinedTypes map[string]struct{}, workspaceUserDefinedTypes *WorkspaceUserDefinedTypeIndex, workspaceUserDefinedTypesComplete bool) bool {
	if workspaceUserDefinedTypesComplete || !inferred.IsArray || !param.IsArray || !isByRefPointerSizedType(param.Type) {
		return false
	}
	if byRefArrayReinterpretationWithWorkspace(inferred, param, localUserDefinedTypes, workspaceUserDefinedTypes) {
		return false
	}
	return a.byRefMayBeUserDefinedType(inferred.Type)
}

func (a Analyzer) byRefMayBeUserDefinedType(typ string) bool {
	typeName := byRefLocalTypeIdentity(typ)
	if typeName == "" || isBuiltinLocalTypeName(byRefShortLocalTypeName(typeName)) {
		return false
	}
	if a.DB != nil {
		if _, ok := a.DB.ResolveType(typeName); ok {
			return false
		}
		if canonical := byRefCanonicalType(typeName); canonical != typeName {
			if _, ok := a.DB.ResolveType(canonical); ok {
				return false
			}
		}
	}
	return true
}

func (a Analyzer) byRefLocalUserDefinedTypesForDocument(doc Document, symbols []Symbol) map[string]struct{} {
	types := make(map[string]struct{})
	for _, symbol := range symbols {
		if symbol.Parent != "" || !strings.EqualFold(strings.TrimSpace(symbol.Kind), "type") || len(symbol.ConditionalBranches) > 0 {
			continue
		}
		if !a.isCurrentModuleSymbol(doc, symbol) && strings.EqualFold(strings.TrimSpace(symbol.Visibility), "Private") {
			continue
		}
		name := byRefShortLocalTypeName(symbol.Name)
		if name == "" {
			continue
		}
		types[name] = struct{}{}
		if module := byRefLocalTypeIdentity(symbol.Module); module != "" {
			types[module+"."+name] = struct{}{}
		}
		if project := byRefLocalTypeIdentity(a.Config.Project.Name); project != "" {
			types[project+"."+name] = struct{}{}
		}
	}
	return types
}

// WorkspaceUserDefinedTypeIndex stores only the project UDT identities that
// are visible outside their declaring module and are unconditional. Private
// types in the current module come from the document-local symbol set instead.
type WorkspaceUserDefinedTypeIndex struct {
	unqualified map[string]struct{}
	qualified   map[string]struct{}
}

// NewWorkspaceUserDefinedTypeIndex builds a names-only index from one
// immutable workspace symbol snapshot. The input is expected to contain type
// symbols, but filtering here keeps custom providers fail-closed.
func NewWorkspaceUserDefinedTypeIndex(symbols []Symbol) *WorkspaceUserDefinedTypeIndex {
	return NewWorkspaceUserDefinedTypeIndexForProject(symbols, "")
}

// NewWorkspaceUserDefinedTypeIndexForProject also registers the VBA project
// name as a qualifier. VBA accepts both Module.TypeName and
// ProjectName.TypeName for project-local user-defined types.
func NewWorkspaceUserDefinedTypeIndexForProject(symbols []Symbol, projectName string) *WorkspaceUserDefinedTypeIndex {
	index := &WorkspaceUserDefinedTypeIndex{
		unqualified: make(map[string]struct{}),
		qualified:   make(map[string]struct{}),
	}
	projectName = byRefLocalTypeIdentity(projectName)
	for _, symbol := range symbols {
		if symbol.Parent != "" || !strings.EqualFold(strings.TrimSpace(symbol.Kind), "type") || len(symbol.ConditionalBranches) > 0 || strings.EqualFold(strings.TrimSpace(symbol.Visibility), "Private") {
			continue
		}
		name := byRefShortLocalTypeName(symbol.Name)
		if name == "" {
			continue
		}
		index.unqualified[name] = struct{}{}
		if module := byRefLocalTypeIdentity(symbol.Module); module != "" {
			index.qualified[module+"."+name] = struct{}{}
		}
		if projectName != "" {
			index.qualified[projectName+"."+name] = struct{}{}
		}
	}
	return index
}

func (index *WorkspaceUserDefinedTypeIndex) matches(typ string) bool {
	if index == nil {
		return false
	}
	typeName := byRefLocalTypeIdentity(typ)
	if strings.Contains(typeName, ".") {
		_, ok := index.qualified[typeName]
		return ok
	}
	_, ok := index.unqualified[byRefShortLocalTypeName(typeName)]
	return ok
}

func byRefLocalTypeIdentity(typ string) string {
	return strings.ToLower(strings.TrimSpace(typ))
}

func byRefShortLocalTypeName(typ string) string {
	typ = byRefLocalTypeIdentity(typ)
	if separator := strings.LastIndex(typ, "."); separator >= 0 {
		typ = typ[separator+1:]
	}
	return typ
}

func isByRefPointerSizedType(typ string) bool {
	// LONG_PTR is the Windows spelling used by VBA API declarations. Treat
	// this explicit alias as equivalent to LongPtr, but do not normalize
	// arbitrary underscores in user-defined type names.
	if byRefCanonicalType(typ) == "longptr" {
		return true
	}
	return strings.EqualFold(strings.TrimSpace(typ), "LONG"+"_"+"PTR")
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
