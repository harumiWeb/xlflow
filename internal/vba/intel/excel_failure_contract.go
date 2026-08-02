package intel

import "strings"

// ExcelAPIFailureContract describes how an Excel member signals that a result
// is unavailable. It intentionally models the member contract, rather than
// asserting that a particular invocation will fail.
type ExcelAPIFailureContract string

const (
	// ExcelAPIFailureRaisesError is used by Excel APIs that raise a VBA runtime
	// error when no result is available.
	ExcelAPIFailureRaisesError ExcelAPIFailureContract = "raises_error"
	// ExcelAPIFailureReturnsErrorValue is used by Excel APIs that return a
	// Variant/Error value which callers must inspect before consuming.
	ExcelAPIFailureReturnsErrorValue ExcelAPIFailureContract = "returns_error_value"
)

// ExcelAPIFailureCall is a resolved call to an Excel API with an exceptional
// failure contract. DirectIsError is true only for a direct IsError(call)
// guard; broader control-flow and assigned-value handling intentionally stays
// with the analyzer layer that has CFG information.
type ExcelAPIFailureCall struct {
	API           string
	Contract      ExcelAPIFailureContract
	Range         Range
	DirectIsError bool
}

// ResolvedExcelAPIFailureCalls finds calls to selected Excel members by their
// resolved receiver type and member name. It deliberately does not infer APIs
// from source text, so late-bound Object members and unrelated project methods
// with the same name are ignored.
func (a Analyzer) ResolvedExcelAPIFailureCalls(doc Document) []ExcelAPIFailureCall {
	if a.DB == nil {
		return nil
	}
	index, ok := a.documentIndexFor(doc)
	if !ok || index == nil {
		return nil
	}
	typeContext := newDocumentTypeContext(doc, documentLines(doc), nil, index)
	out := make([]ExcelAPIFailureCall, 0)
	for _, logicalLine := range logicalLinesForCallAnalysis(doc.Source) {
		calls := callsOnLine(logicalLine.Text)
		for _, call := range calls {
			if !excelAPIFailureCallCandidate(call.Target) {
				continue
			}
			callRange := logicalLine.callRange(call)
			sig, resolved, err := a.resolveCallSignatureAtContext(doc, call.Target, callRange.Start, []Document{doc}, typeContext)
			if err != nil || !resolved {
				continue
			}
			contract, api, found := excelAPIFailureContract(sig)
			if !found {
				continue
			}
			if contract == ExcelAPIFailureReturnsErrorValue && discardedCallResult(call) {
				continue
			}
			out = append(out, ExcelAPIFailureCall{
				API:           api,
				Contract:      contract,
				Range:         callRange,
				DirectIsError: contract == ExcelAPIFailureReturnsErrorValue && a.directIsErrorGuard(doc, logicalLine, calls, call, typeContext),
			})
		}
	}
	return out
}

// excelAPIFailureCallCandidate avoids resolving every VBA call in a document.
// The failure-contract table is intentionally closed: only these member names
// can produce VBA218, so resolving any other call cannot change its result.
// Receiver type resolution remains authoritative for the retained candidates.
func excelAPIFailureCallCandidate(target string) bool {
	_, member, hasReceiver := splitCallTarget(target)
	if !hasReceiver {
		member = strings.TrimSpace(strings.TrimPrefix(strings.TrimSpace(target), "."))
	}
	switch strings.ToLower(member) {
	case "specialcells", "match", "vlookup", "xlookup", "index":
		return true
	default:
		return false
	}
}

// discardedCallResult excludes a standalone function call: its Variant/Error
// result is not consumed, so it cannot cause a later value-use failure.
func discardedCallResult(call parsedCall) bool {
	line := strings.TrimSpace(call.Line)
	candidate := strings.TrimSpace(call.Line[call.Start:call.End])
	line = strings.TrimSpace(strings.TrimPrefix(strings.ToLower(line), "call "))
	return strings.EqualFold(line, candidate)
}

// ExcelAPIFailureContractDiagnostics adapts typed Excel failure-contract calls
// to VBA218 diagnostics. Callers that have control-flow information can use
// ResolvedExcelAPIFailureCalls directly to apply additional safe-path guards.
func (a Analyzer) ExcelAPIFailureContractDiagnostics(doc Document) []Diagnostic {
	calls := a.ResolvedExcelAPIFailureCalls(doc)
	out := make([]Diagnostic, 0, len(calls))
	for _, call := range calls {
		if call.Contract == ExcelAPIFailureReturnsErrorValue && call.DirectIsError {
			continue
		}
		out = append(out, Diagnostic{
			Code:       "VBA218",
			Severity:   "warning",
			Source:     "xlflow",
			Message:    excelAPIFailureContractMessage(call),
			Range:      call.Range,
			Rule:       "VBA218",
			Confidence: "high",
		})
	}
	return out
}

func excelAPIFailureContract(sig Signature) (ExcelAPIFailureContract, string, bool) {
	receiver := strings.TrimSpace(sig.receiverType)
	member := strings.TrimSpace(sig.memberName)
	switch {
	case strings.EqualFold(receiver, "Excel.Range") && strings.EqualFold(member, "SpecialCells"):
		return ExcelAPIFailureRaisesError, "Range.SpecialCells", true
	case strings.EqualFold(receiver, "Excel.WorksheetFunction") && oneOfFold(member, "Match", "VLookup", "XLookup", "Index"):
		return ExcelAPIFailureRaisesError, "WorksheetFunction." + member, true
	case strings.EqualFold(receiver, "Excel.Application") && oneOfFold(member, "Match", "VLookup", "XLookup"):
		return ExcelAPIFailureReturnsErrorValue, "Application." + member, true
	default:
		return "", "", false
	}
}

func excelAPIFailureContractMessage(call ExcelAPIFailureCall) string {
	if call.Contract == ExcelAPIFailureReturnsErrorValue {
		return call.API + " may return a Variant/Error when no result is available; check it with IsError before consuming the value."
	}
	return call.API + " may raise a runtime error when no result is available; handle that error path locally."
}

func oneOfFold(value string, candidates ...string) bool {
	for _, candidate := range candidates {
		if strings.EqualFold(value, candidate) {
			return true
		}
	}
	return false
}

func (a Analyzer) directIsErrorGuard(doc Document, logicalLine logicalCallAnalysisLine, calls []parsedCall, candidate parsedCall, typeContext *documentTypeContext) bool {
	candidateText := strings.TrimSpace(candidate.Line[candidate.Start:candidate.End])
	if candidateText == "" {
		return false
	}
	for _, outer := range calls {
		if outer.Start >= candidate.Start || outer.End <= candidate.End || len(outer.Arguments) != 1 {
			continue
		}
		outerRange := logicalLine.callRange(outer)
		sig, resolved, err := a.resolveCallSignatureAtContext(doc, outer.Target, outerRange.Start, []Document{doc}, typeContext)
		if err != nil || !resolved || !strings.EqualFold(sig.receiverType, "VBA.Global") || !strings.EqualFold(sig.memberName, "IsError") {
			continue
		}
		if strings.EqualFold(strings.TrimSpace(outer.Arguments[0].Text), candidateText) {
			return true
		}
	}
	return false
}
