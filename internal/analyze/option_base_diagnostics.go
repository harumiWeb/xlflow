package analyze

import (
	"fmt"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// optionBaseInconsistencyFindings implements VBA271 and VBA272, the opt-in
// Option Base consistency diagnostics. Both rules report only inside a module
// that explicitly declares `Option Base 1`: `Option Base 0` and the default
// (no Option Base directive) already agree with the zero-based constructs, so
// they stay silent. A recovered or multi-branch procedure symbol is skipped
// because its parameter/call projection may be incomplete.
func (a Analyzer) optionBaseInconsistencyFindings(file parsedFile, proc sourceProcedure) []Finding {
	// Both rules are opt-in; check the flags before touching module facts so
	// the common disabled configuration costs one branch per procedure.
	if !a.Config.Analyze.DetectOptionBaseArrayInconsistency && !a.Config.Analyze.DetectOptionBaseParamArrayInconsistency {
		return nil
	}
	if proc.IR == nil || proc.IR.Symbol.Recovered || len(proc.IR.Symbol.ConditionalBranches) > 0 {
		return nil
	}
	if arrayOptionBase(file) != 1 {
		return nil
	}
	var findings []Finding
	if a.Config.Analyze.DetectOptionBaseArrayInconsistency {
		findings = append(findings, a.optionBaseArrayFindings(file, proc)...)
	}
	if a.Config.Analyze.DetectOptionBaseParamArrayInconsistency {
		findings = append(findings, a.optionBaseParamArrayFindings(file, proc)...)
	}
	return findings
}

// optionBaseArrayFindings implements VBA272: an explicitly qualified
// `VBA.Array(...)` call is reported because the type-library-qualified
// intrinsic always returns a zero-based array regardless of the module's
// Option Base. An unqualified `Array(...)` call is never reported: per the
// language reference it honors Option Base, so under `Option Base 1` it
// already returns a one-based array — consistent with the module's declared
// base. A qualified call on any other receiver (`foo.Array(...)`) may be a
// project or host member returning an arbitrary lower bound and stays
// silent. The `VBA.` qualifier names the type library itself and cannot be
// shadowed, so no call resolution is needed.
func (a Analyzer) optionBaseArrayFindings(file parsedFile, proc sourceProcedure) []Finding {
	var findings []Finding
	for call := range proc.Calls.All() {
		if call.IsRaiseEvent || !strings.EqualFold(call.Callee.BaseName, "array") {
			continue
		}
		// Indexed assignment targets such as `VBA.Array(0) = "a"` are
		// recorded as call-shaped facts but are never invocations.
		if procedureir.IsAssignmentTargetCall(call, *proc.IR) {
			continue
		}
		if call.Callee.Receiver == nil || !strings.EqualFold(cleanIdentifier(*call.Callee.Receiver), "vba") {
			continue
		}
		finding := a.simpleFinding(
			file, proc, call.Range.StartLine, "VBA272", "warning",
			fmt.Sprintf("%s returns a zero-based array even though this module declares Option Base 1.", strings.TrimSpace(call.Callee.Text)),
			"VBA.Array always produces a lower bound of 0; Option Base only changes the default bound of Dim/ReDim declarations and unqualified Array calls.",
			"Review indexing assumptions on the returned array, or construct it with explicit bounds instead.",
		)
		finding.Column = call.Range.StartColumn
		finding.EndLine = call.Range.EndLine
		finding.EndColumn = call.Range.EndColumn
		finding.ScopeEndLine = proc.EndLine
		findings = append(findings, finding)
	}
	return findings
}

// optionBaseParamArrayFindings implements VBA271: a ParamArray parameter is
// reported because the host always passes it as a zero-based Variant array
// regardless of the module's Option Base.
func (a Analyzer) optionBaseParamArrayFindings(file parsedFile, proc sourceProcedure) []Finding {
	var findings []Finding
	for parameter := range proc.Params.All() {
		if !parameter.ParamArray {
			continue
		}
		// Anchor on the parameter identifier, not the whole `ParamArray ...`
		// declaration; fall back to the declaration span only when the parser
		// could not resolve a name token.
		anchor := parameter.Range
		if parameter.NameRange != nil {
			anchor = *parameter.NameRange
		}
		line := anchor.StartLine
		if line < 1 {
			line = proc.IR.Symbol.DeclarationRange.StartLine
		}
		finding := a.simpleFinding(
			file, proc, line, "VBA271", "warning",
			fmt.Sprintf("ParamArray '%s' is always zero-based even though this module declares Option Base 1.", cleanIdentifier(parameter.Name)),
			"The host supplies ParamArray arguments as a zero-based Variant array; Option Base does not change its lower bound.",
			"Index the parameter from 0, or document that this signature intentionally mixes array bases.",
		)
		finding.Column = anchor.StartColumn
		finding.EndLine = anchor.EndLine
		finding.EndColumn = anchor.EndColumn
		finding.ScopeEndLine = proc.EndLine
		findings = append(findings, finding)
	}
	return findings
}
