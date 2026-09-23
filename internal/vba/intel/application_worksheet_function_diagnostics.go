package intel

import (
	"context"
	"fmt"
	"strings"
)

const (
	applicationTypeName          = "Excel.Application"
	worksheetFunctionTypeName    = "Excel.WorksheetFunction"
	applicationWorksheetRuleCode = "VBA261"
)

// ApplicationWorksheetFunctionDispatchCall describes an explicitly typed
// Application call whose member is exposed by Excel.WorksheetFunction.
// Application dispatch remains late-bound even when the Application receiver
// itself has a generated Excel TypeLib type.
type ApplicationWorksheetFunctionDispatchCall struct {
	Member     string
	ReturnType string
	Range      Range
}

// ResolvedApplicationWorksheetFunctionDispatchCalls finds explicit calls on
// an Excel.Application receiver whose member belongs to the complete
// generated Excel.WorksheetFunction member set. Unknown, late-bound, and
// incomplete TypeLib cases remain unresolved and fail open.
func (a Analyzer) ResolvedApplicationWorksheetFunctionDispatchCalls(doc Document) []ApplicationWorksheetFunctionDispatchCall {
	out, _ := a.ResolvedApplicationWorksheetFunctionDispatchCallsContext(context.Background(), doc)
	return out
}

// ResolvedApplicationWorksheetFunctionDispatchCallsContext is the cancellable
// form used by batch and realtime analysis.
func (a Analyzer) ResolvedApplicationWorksheetFunctionDispatchCallsContext(ctx context.Context, doc Document) ([]ApplicationWorksheetFunctionDispatchCall, error) {
	if ctx == nil {
		ctx = context.Background()
	}
	if a.DB == nil || a.TypeDBResolutionIncomplete || !a.completeMemberSetType(worksheetFunctionTypeName) {
		return nil, nil
	}
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	index, ok := a.documentIndexFor(doc)
	if !ok || index == nil {
		return nil, nil
	}
	typeContext := newDocumentTypeContext(doc, documentLines(doc), nil, index)
	out := make([]ApplicationWorksheetFunctionDispatchCall, 0)
	for i, logicalLine := range logicalLinesForCallAnalysis(doc.Source) {
		if i&0x3f == 0 {
			if err := ctx.Err(); err != nil {
				return nil, err
			}
		}
		for _, call := range callsOnLine(logicalLine.Text) {
			_, member, qualified := splitCallTarget(call.Target)
			if !qualified && !strings.HasPrefix(strings.TrimSpace(call.Target), ".") {
				continue
			}
			if member == "" {
				continue
			}
			worksheetMember, found := a.DB.ResolveMember(worksheetFunctionTypeName, member)
			if !found {
				continue
			}
			callRange := logicalLine.callRange(call)
			receiverType, resolved := a.typedMemberReceiverType(doc, call.Target, callRange.Start, byteOffsetForDocumentPosition(doc, callRange.Start), typeContext)
			if !resolved || !strings.EqualFold(receiverType, applicationTypeName) {
				continue
			}
			out = append(out, ApplicationWorksheetFunctionDispatchCall{
				Member:     worksheetMember.Name,
				ReturnType: worksheetMember.ReturnType,
				Range:      logicalCallMemberRange(logicalLine, call, member),
			})
		}
	}
	return out, ctx.Err()
}

// ApplicationWorksheetFunctionDispatchDiagnostics reports calls that use
// Application worksheet-function dispatch instead of the strongly typed
// Application.WorksheetFunction form.
func (a Analyzer) ApplicationWorksheetFunctionDispatchDiagnostics(doc Document) []Diagnostic {
	out, _ := a.ApplicationWorksheetFunctionDispatchDiagnosticsContext(context.Background(), doc)
	return out
}

// ApplicationWorksheetFunctionDispatchDiagnosticsContext adapts resolved
// Application worksheet-function calls to VBA261 diagnostics. It intentionally
// leaves failure-contract ownership to VBA218 and approximate-lookup ownership
// to VBA251.
func (a Analyzer) ApplicationWorksheetFunctionDispatchDiagnosticsContext(ctx context.Context, doc Document) ([]Diagnostic, error) {
	if ctx == nil {
		ctx = context.Background()
	}
	calls, err := a.ResolvedApplicationWorksheetFunctionDispatchCallsContext(ctx, doc)
	if err != nil {
		return nil, err
	}
	out := make([]Diagnostic, 0, len(calls))
	for i, call := range calls {
		if i&0x3f == 0 {
			if err := ctx.Err(); err != nil {
				return nil, err
			}
		}
		out = append(out, Diagnostic{
			Code:       applicationWorksheetRuleCode,
			Severity:   "information",
			Source:     "xlflow",
			Message:    applicationWorksheetFunctionDispatchMessage(call.Member, call.ReturnType),
			Range:      call.Range,
			Rule:       applicationWorksheetRuleCode,
			Confidence: "high",
		})
	}
	return out, ctx.Err()
}

func applicationWorksheetFunctionDispatchMessage(member, returnType string) string {
	return fmt.Sprintf("Application.%s uses late-bound worksheet-function dispatch, which can return worksheet error values with Variant semantics; Application.WorksheetFunction.%s is strongly typed as %s by the generated Excel TypeLib and can raise a runtime error for worksheet errors.", member, member, firstNonEmpty(strings.TrimSpace(returnType), "Variant"))
}
