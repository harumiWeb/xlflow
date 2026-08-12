package procedureir

import (
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
)

// ResolutionDiagnostic is the protocol-neutral projection of a deterministic
// compile-equivalent resolution failure.  Batch, LSP, and lint adapters attach
// their own file/procedure metadata while preserving this range and message.
type ResolutionDiagnostic struct {
	Code       string
	Message    string
	Range      vbaast.Range
	Candidates []Candidate
}

// ResolutionSuggestion returns the canonical remediation text for a
// deterministic project-resolution diagnostic.  Batch analysis, lint, and
// editor adapters share this lookup so the same VB052-VB054 finding does not
// drift across surfaces.
func ResolutionSuggestion(code string) string {
	switch strings.TrimSpace(code) {
	case "VB052":
		return "Call a declared Sub, Function, or Property, or qualify the project-local target correctly."
	case "VB053":
		return "Qualify the Enum member with its Enum name."
	case "VB054":
		return "Declare the Event in this object module before raising it."
	default:
		return ""
	}
}

// Diagnostics returns only negative outcomes that are provable from this
// resolved document.  Incomplete snapshots, parser recovery, external/member
// calls, and dynamic APIs deliberately fail open.
func Diagnostics(doc DocumentIR, complete bool) []ResolutionDiagnostic {
	if !complete || documentResolutionIncomplete(doc) {
		return nil
	}
	out := make([]ResolutionDiagnostic, 0)
	for _, procedure := range doc.Procedures {
		if procedure.Symbol.Recovered {
			continue
		}
		for _, call := range procedure.Calls {
			if call.IsRaiseEvent {
				continue
			}
			if !syntacticInvocation(procedure, call) {
				continue
			}
			if call.Resolution.Status == ResolutionNonCallable ||
				(call.Resolution.Status == ResolutionUnresolved && call.Resolution.ProjectLocal) {
				name := strings.TrimSpace(call.Callee.Text)
				message := "Call target is missing or is not callable in this project."
				if name != "" {
					message = "Call target " + name + " is missing or is not callable in this project."
				}
				out = append(out, ResolutionDiagnostic{Code: "VB052", Message: message, Range: call.Range, Candidates: append([]Candidate(nil), call.Resolution.Candidates...)})
			}
		}
		for _, event := range procedure.RaiseEvents {
			if event.Recovered || event.Resolution.Status != ResolutionUnresolved {
				continue
			}
			out = append(out, ResolutionDiagnostic{Code: "VB054", Message: "RaiseEvent target is not declared in this object module.", Range: event.Range})
		}
		for _, access := range procedure.Accesses {
			if accessIsQualified(procedure, access) {
				continue
			}
			if access.Resolution.Status != ResolutionAmbiguous || len(access.Resolution.Candidates) < 2 {
				continue
			}
			allEnum := true
			for _, candidate := range access.Resolution.Candidates {
				if !isEnumMemberKind(candidate.Kind) {
					allEnum = false
					break
				}
			}
			if allEnum {
				out = append(out, ResolutionDiagnostic{Code: "VB053", Message: "Enum member reference is ambiguous; qualify the member with its Enum name.", Range: access.Range, Candidates: append([]Candidate(nil), access.Resolution.Candidates...)})
			}
		}
	}
	return out
}

func accessIsQualified(procedure ProcedureIR, access VariableAccess) bool {
	if access.ExpressionID <= 0 {
		// A missing expression link means the qualifier cannot be established
		// from syntax.  Suppress the negative diagnostic rather than treating
		// incomplete IR as proof of a bare reference.
		return true
	}
	for _, expression := range procedure.Expressions {
		if expression.ID != access.ExpressionID {
			continue
		}
		text := strings.TrimSpace(expression.Text)
		if text == "" {
			return true
		}
		return strings.ContainsAny(text, ".!")
	}
	// The expression ID may refer to a node omitted by a recovered parse.
	// Fail open when the link cannot be resolved.
	return true
}

// syntacticInvocation filters parser call facts that are actually scalar
// expression reads.  VBA's CST represents some expression statements (for
// example `i + 1 = limit` and an enum argument passed to another call) as a
// zero-receiver call.  Those facts cannot establish a non-callable target.
func syntacticInvocation(procedure ProcedureIR, call CallSite) bool {
	if call.StatementID <= 0 || call.StatementID > len(procedure.Statements) {
		return true
	}
	statement := procedure.Statements[call.StatementID-1]
	text := strings.TrimSpace(statement.Text)
	callee := strings.TrimSpace(call.Callee.Text)
	if text == "" || callee == "" {
		return true
	}
	lowerText := strings.ToLower(text)
	idx := indexTokenOccurrence(lowerText, strings.ToLower(callee))
	if idx < 0 {
		// The parser supplied a call fact whose callee cannot be found in the
		// statement.  There is no syntactic evidence to project as a negative
		// diagnostic, so fail open.
		return false
	}
	rest := strings.TrimSpace(text[idx+len(callee):])
	if rest == "" {
		if statement.Kind == StatementAssignment || statement.Kind == StatementSet {
			// A bare identifier on the right-hand side is an expression read;
			// without parentheses there is no syntactic evidence of invocation.
			return false
		}
		// Nested identifier reads in a call statement have an expression ID
		// but no argument list/parentheses (e.g. ADO_ASYNC_OPTION).
		return call.ExpressionID == 0 || strings.Contains(text, "(")
	}
	first := rest[0]
	return !strings.ContainsRune("+-*/=<>:&", rune(first))
}

// indexTokenOccurrence finds a case-insensitive callee occurrence without
// matching it as a substring of a larger VBA identifier (for example the
// `Helper` suffix in `MyHelperValue`).
func indexTokenOccurrence(text, token string) int {
	if token == "" {
		return -1
	}
	for start := 0; start <= len(text)-len(token); {
		relative := strings.Index(text[start:], token)
		if relative < 0 {
			return -1
		}
		idx := start + relative
		beforeOK := idx == 0 || !vbaIdentifierByte(text[idx-1])
		after := idx + len(token)
		afterOK := after >= len(text) || !vbaIdentifierByte(text[after])
		if beforeOK && afterOK {
			return idx
		}
		start = idx + 1
	}
	return -1
}

func vbaIdentifierByte(value byte) bool {
	return value == '_' || value == '$' || value == '%' || value == '&' || value == '@' || value == '^' || value == '!' ||
		(value >= '0' && value <= '9') || (value >= 'a' && value <= 'z') || (value >= 'A' && value <= 'Z')
}

func documentResolutionIncomplete(doc DocumentIR) bool {
	if doc.Parse.HasError || doc.Parse.HasMissing {
		return true
	}
	for _, declaration := range doc.Declarations {
		if declaration.Recovered || len(declaration.ConditionalBranches) > 0 {
			return true
		}
	}
	for _, procedure := range doc.Procedures {
		if procedure.Symbol.Recovered || len(procedure.Symbol.ConditionalBranches) > 0 {
			return true
		}
		for _, declaration := range procedure.Declarations {
			if declaration.Recovered || len(declaration.ConditionalBranches) > 0 {
				return true
			}
		}
		for _, statement := range procedure.Statements {
			if statement.Recovered {
				return true
			}
		}
		for _, event := range procedure.RaiseEvents {
			if event.Recovered || len(event.ConditionalBranches) > 0 {
				return true
			}
		}
	}
	return false
}
