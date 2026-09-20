package intel

import (
	"context"
	"fmt"
	"strings"
)

// TypedMemberDiagnosticSpec describes a diagnostic that is emitted when a
// statically typed receiver has a complete member set but does not expose the
// called member. The resolver intentionally accepts only generated TypeLib
// metadata for complete sets; curated and incomplete metadata remain
// fail-open.
type TypedMemberDiagnosticSpec struct {
	Code       string
	Severity   string
	Rule       string
	Confidence string
	Receiver   string
	Message    func(member string) string
}

// TypedMemberDiagnostics reports calls to unavailable members on a receiver
// whose generated TypeLib member set is complete. It shares the document/type
// resolution path used by signatures, completion, and existing member
// diagnostics, including local declarations and With blocks.
func (a Analyzer) TypedMemberDiagnostics(doc Document, spec TypedMemberDiagnosticSpec) []Diagnostic {
	out, _ := a.TypedMemberDiagnosticsContext(context.Background(), doc, spec)
	return out
}

// TypedMemberDiagnosticsContext is the cancellable form used by batch and
// realtime analysis.
func (a Analyzer) TypedMemberDiagnosticsContext(ctx context.Context, doc Document, spec TypedMemberDiagnosticSpec) ([]Diagnostic, error) {
	if ctx == nil {
		ctx = context.Background()
	}
	if a.DB == nil || a.TypeDBResolutionIncomplete || strings.TrimSpace(spec.Receiver) == "" {
		return nil, nil
	}
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	index, ok := a.documentIndexFor(doc)
	if !ok || index == nil {
		return nil, nil
	}
	if !a.completeMemberSetType(spec.Receiver) {
		return nil, nil
	}
	typeContext := newDocumentTypeContext(doc, documentLines(doc), nil, index)
	code := firstNonEmpty(spec.Code, "VBA252")
	severity := firstNonEmpty(spec.Severity, "warning")
	rule := firstNonEmpty(spec.Rule, code)
	confidence := firstNonEmpty(spec.Confidence, "high")
	message := spec.Message
	if message == nil {
		message = func(member string) string {
			return fmt.Sprintf("%s does not expose member %q in its generated TypeLib.", spec.Receiver, member)
		}
	}

	var out []Diagnostic
	for i, logicalLine := range logicalLinesForCallAnalysis(doc.Source) {
		if i&0x3f == 0 {
			if err := ctx.Err(); err != nil {
				return nil, err
			}
		}
		for _, call := range callsOnLine(logicalLine.Text) {
			_, memberName, qualified := splitCallTarget(call.Target)
			if !qualified && strings.HasPrefix(strings.TrimSpace(call.Target), ".") {
				memberName = strings.TrimSpace(strings.TrimPrefix(strings.TrimSpace(call.Target), "."))
			}
			if memberName == "" || (!qualified && !strings.HasPrefix(strings.TrimSpace(call.Target), ".")) {
				continue
			}
			callRange := logicalLine.callRange(call)
			receiverType, resolved := a.typedMemberReceiverType(doc, call.Target, callRange.Start, byteOffsetForDocumentPosition(doc, callRange.Start), typeContext)
			if !resolved || !strings.EqualFold(receiverType, spec.Receiver) {
				continue
			}
			if _, found := a.DB.ResolveMember(receiverType, memberName); found {
				continue
			}
			memberRange := logicalCallMemberRange(logicalLine, call, memberName)
			out = append(out, Diagnostic{
				Code:       code,
				Severity:   severity,
				Source:     "xlflow",
				Message:    message(memberName),
				Range:      memberRange,
				Rule:       rule,
				Confidence: confidence,
			})
		}
	}
	return out, ctx.Err()
}

// UnavailableWorksheetFunctionMemberDiagnostics reports calls to members that
// are absent from a generated Excel.WorksheetFunction TypeLib type.
func (a Analyzer) UnavailableWorksheetFunctionMemberDiagnostics(doc Document) []Diagnostic {
	out, _ := a.UnavailableWorksheetFunctionMemberDiagnosticsContext(context.Background(), doc)
	return out
}

// UnavailableWorksheetFunctionMemberDiagnosticsContext is the cancellable
// WorksheetFunction-specific typed-member diagnostic entry point.
func (a Analyzer) UnavailableWorksheetFunctionMemberDiagnosticsContext(ctx context.Context, doc Document) ([]Diagnostic, error) {
	return a.TypedMemberDiagnosticsContext(ctx, doc, TypedMemberDiagnosticSpec{
		Code:       "VBA252",
		Severity:   "warning",
		Rule:       "VBA252",
		Confidence: "high",
		Receiver:   "Excel.WorksheetFunction",
		Message: func(member string) string {
			return fmt.Sprintf("WorksheetFunction member %q is not available in the generated Excel TypeLib.", member)
		},
	})
}

func (a Analyzer) typedMemberReceiverType(doc Document, target string, pos Position, offset int, typeContext *documentTypeContext) (string, bool) {
	canonical := func(typ string) string {
		return canonicalDiagnosticType(a.DB, typ)
	}
	receiverExpr, _, qualified := splitCallTarget(target)
	if strings.HasPrefix(strings.TrimSpace(target), ".") {
		receiverType, ok := a.withBlockTypeAtContext(doc, pos, offset, typeContext)
		if !ok {
			return "", false
		}
		if !qualified || strings.TrimSpace(receiverExpr) == "" || strings.TrimSpace(receiverExpr) == "." {
			return canonical(receiverType), true
		}
		resolvedType, ok := a.resolveRelativeMemberExpressionType(receiverType, receiverExpr)
		if !ok {
			return "", false
		}
		return canonical(resolvedType), true
	}
	if !qualified {
		return "", false
	}
	resolvedType, ok := a.resolveDocumentExpressionTypeAtContext(doc, receiverExpr, offset, typeContext)
	if !ok {
		return "", false
	}
	return canonical(resolvedType), true
}

func logicalCallMemberRange(line logicalCallAnalysisLine, call parsedCall, member string) Range {
	start := call.Start
	if index := strings.LastIndex(strings.ToLower(call.Target), strings.ToLower(member)); index >= 0 {
		start += index
	}
	startLine, startColumn := line.positionForOffset(start)
	endLine, endColumn := line.positionForOffset(start + len(member))
	return Range{
		Start: Position{Line: startLine, Character: startColumn},
		End:   Position{Line: endLine, Character: endColumn},
	}
}
