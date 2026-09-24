package analyze

import (
	"fmt"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// optionBaseInconsistencyFindings implements VBA270 and VBA271, the opt-in
// Option Base consistency diagnostics. Both rules report only inside a module
// that explicitly declares `Option Base 1`: `Option Base 0` and the default
// (no Option Base directive) already agree with the zero-based constructs, so
// they stay silent. A recovered or multi-branch procedure symbol is skipped
// because its parameter/call projection may be incomplete.
func (a Analyzer) optionBaseInconsistencyFindings(file parsedFile, proc sourceProcedure, resolver procedureir.Resolver) []Finding {
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
		findings = append(findings, a.optionBaseArrayFindings(file, proc, resolver)...)
	}
	if a.Config.Analyze.DetectOptionBaseParamArrayInconsistency {
		findings = append(findings, a.optionBaseParamArrayFindings(file, proc)...)
	}
	return findings
}

// optionBaseArrayFindings implements VBA270: an `Array(...)` call that provably
// resolves to the VBA intrinsic is reported because `VBA.Array` always returns
// a zero-based array regardless of the module's Option Base. Only
// ResolutionBuiltinLike counts: a user-defined `Array` procedure
// (ResolutionMatched), a lexical non-callable shadow such as `Dim Array`
// (ResolutionNonCallable), and every ambiguous, member, external, unresolved,
// or dynamic outcome stay silent.
func (a Analyzer) optionBaseArrayFindings(file parsedFile, proc sourceProcedure, resolver procedureir.Resolver) []Finding {
	var findings []Finding
	// The shadow verdict is loop-invariant; compute it lazily on the first
	// unqualified Array candidate instead of rescanning declarations per call.
	shadowChecked, shadowed := false, false
	for call := range proc.Calls.All() {
		if call.IsRaiseEvent || !strings.EqualFold(call.Callee.BaseName, "array") {
			continue
		}
		// Indexed assignment targets such as `Array(0) = "a"` are recorded as
		// call-shaped facts but are never invocations: they bind to a lexical
		// array variable, and the resolver deliberately leaves their
		// NonCallableNames empty.
		if procedureir.IsAssignmentTargetCall(call, *proc.IR) {
			continue
		}
		// A qualified call is evidence only for the explicit `VBA.Array` form:
		// `foo.Array(...)` on an unknown receiver may be a project or host
		// member returning an arbitrary lower bound.
		if call.Callee.Receiver != nil && !strings.EqualFold(cleanIdentifier(*call.Callee.Receiver), "vba") {
			continue
		}
		// The resolver does not list array declarations in NonCallableNames
		// because an indexed read `arr(i)` is grammar-identical to a call.
		// Any in-scope declaration named `Array` therefore still needs this
		// lexical shadow check: under VBA rules it wins over the intrinsic for
		// an unqualified reference, whatever its declared shape.
		if call.Callee.Receiver == nil {
			if !shadowChecked {
				shadowed = optionBaseArrayShadowed(file, proc)
				shadowChecked = true
			}
			if shadowed {
				continue
			}
		}
		resolution := call.Resolution
		if resolution.Status == procedureir.ResolutionNotAttempted && resolver != nil {
			resolution = resolver.ResolveCall(call)
		}
		if resolution.Status != procedureir.ResolutionBuiltinLike {
			continue
		}
		finding := a.simpleFinding(
			file, proc, call.Range.StartLine, "VBA270", "warning",
			fmt.Sprintf("%s returns a zero-based array even though this module declares Option Base 1.", strings.TrimSpace(call.Callee.Text)),
			"VBA.Array always produces a lower bound of 0; Option Base only changes the default bound of Dim/ReDim declarations.",
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

// optionBaseArrayShadowed reports whether any in-scope declaration is named
// `Array`: procedure locals and parameters first, then the module-level
// declaration projection. Module declarations are keyed by lowercase name.
// A project procedure named `Array` in another module is not a shadow here;
// it is already excluded by the call resolution status.
func optionBaseArrayShadowed(file parsedFile, proc sourceProcedure) bool {
	for declaration := range proc.Declarations.All() {
		if strings.EqualFold(cleanIdentifier(declaration.Name), "array") {
			return true
		}
	}
	for parameter := range proc.Params.All() {
		if strings.EqualFold(cleanIdentifier(parameter.Name), "array") {
			return true
		}
	}
	_, shadowed := file.moduleDecls()["array"]
	return shadowed
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
