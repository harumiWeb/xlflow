package analyze

import (
	"bytes"
	"context"
	"fmt"
	"sort"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/calls"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
)

// callSiteDiscardsReturn reports whether a call is the complete statement, so
// the callee's return value is produced and immediately dropped. Call
// expressions nested in arguments, conditions, or assignment values occupy a
// strictly smaller range than their statement and never satisfy this
// predicate. Indexed assignment targets such as `Foo(x) = v` keep
// StatementAssignment kind and are excluded before the range comparison.
func callSiteDiscardsReturn(call procedureir.CallSite, proc *procedureir.ProcedureIR, source []byte) bool {
	if proc == nil || call.IsRaiseEvent {
		return false
	}
	if call.StatementID <= 0 || call.StatementID > len(proc.Statements) {
		return false
	}
	statement := proc.Statements[call.StatementID-1]
	if statement.Kind != procedureir.StatementCall || statement.SyntaxKind != "call_statement" {
		return false
	}
	if procedureir.IsAssignmentTargetCall(call, *proc) {
		return false
	}
	if call.Range.StartByte != statement.Range.StartByte || call.Range.EndByte != statement.Range.EndByte {
		return false
	}
	return !hasWhitespaceSeparatedPredecessor(proc, statement, source)
}

// hasWhitespaceSeparatedPredecessor reports whether another statement ends on
// the same line immediately before statement with only whitespace between
// them. Valid VBA requires a colon between same-line statements, so a
// whitespace-only gap means the grammar split one logical statement — for
// example a paren-less member call whose member name is a reserved keyword,
// such as `web_Http.Open MethodToName(x), url, async` — into an expression
// statement plus a bogus call statement. Reporting the bogus call would
// misclassify a consumed argument as a discarded result.
func hasWhitespaceSeparatedPredecessor(proc *procedureir.ProcedureIR, statement procedureir.Statement, source []byte) bool {
	if len(source) == 0 {
		return false
	}
	var predecessor *procedureir.Statement
	for i := range proc.Statements {
		candidate := &proc.Statements[i]
		if candidate.ID == statement.ID ||
			candidate.Range.EndLine != statement.Range.StartLine ||
			candidate.Range.EndByte > statement.Range.StartByte {
			continue
		}
		if predecessor == nil || candidate.Range.EndByte > predecessor.Range.EndByte {
			predecessor = candidate
		}
	}
	if predecessor == nil || statement.Range.StartByte > len(source) {
		return false
	}
	// A label_statement (named `Helper :` or numeric `10`) is a legitimate
	// same-line predecessor: its colon or label syntax already separates the
	// following statement.
	if predecessor.SyntaxKind == "label_statement" {
		return false
	}
	for _, b := range source[predecessor.Range.EndByte:statement.Range.StartByte] {
		if b != ' ' && b != '\t' {
			return false
		}
	}
	return true
}

// discardedReturnCandidate resolves a call to a unique project Function or
// Property Get. Unresolved, ambiguous, member, external, and
// non-value-returning callees are not evidence for either rule and stay
// silent (fail-open).
func discardedReturnCandidate(call procedureir.CallSite, resolver procedureir.Resolver) (procedureir.Candidate, bool) {
	resolution := call.Resolution
	if resolution.Status == procedureir.ResolutionNotAttempted && resolver != nil {
		resolution = resolver.ResolveCall(call)
	}
	if resolution.Status != procedureir.ResolutionMatched || len(resolution.Candidates) != 1 {
		return procedureir.Candidate{}, false
	}
	candidate := resolution.Candidates[0]
	switch strings.ToLower(strings.TrimSpace(candidate.Kind)) {
	case string(procedureir.ProcedureFunction), string(procedureir.ProcedurePropertyGet):
		return candidate, true
	default:
		return procedureir.Candidate{}, false
	}
}

// discardedReturnFindings implements VBA257, the opt-in call-site diagnostic
// for resolved Function/Property Get invocations whose result is dropped.
func (a Analyzer) discardedReturnFindings(file parsedFile, proc sourceProcedure, resolver procedureir.Resolver) []Finding {
	if !a.Config.Analyze.DetectDiscardedFunctionReturn || proc.IR == nil {
		return nil
	}
	var findings []Finding
	for call := range proc.Calls.All() {
		if !callSiteDiscardsReturn(call, proc.IR, file.Source) {
			continue
		}
		candidate, ok := discardedReturnCandidate(call, resolver)
		if !ok {
			continue
		}
		qualified := strings.TrimSpace(candidate.QualifiedName)
		if qualified == "" {
			qualified = strings.TrimSpace(call.Callee.Text)
		}
		kindLabel := "Function"
		if strings.EqualFold(candidate.Kind, string(procedureir.ProcedurePropertyGet)) {
			kindLabel = "Property Get"
		}
		finding := a.simpleFinding(
			file, proc, call.Range.StartLine, "VBA257", "warning",
			fmt.Sprintf("%s '%s' returns a value that this call discards.", kindLabel, qualified),
			"The invocation is a standalone call statement, so the returned value is produced and then dropped.",
			"Consume the return value, convert the callee to a Sub, or suppress VBA257 when the discard is intentional.",
		)
		finding.Column = call.Range.StartColumn
		finding.EndLine = call.Range.EndLine
		finding.EndColumn = call.Range.EndColumn
		finding.ScopeEndLine = proc.EndLine
		findings = append(findings, finding)
	}
	return findings
}

// alwaysDiscardedCandidate tracks one VBA258-eligible Function/Property Get
// and the project-wide evidence collected for it.
type alwaysDiscardedCandidate struct {
	fileIndex int
	procIndex int
	discards  int
	consumed  bool
	uncertain bool
}

// functionAlwaysDiscardedFindings implements VBA258, the opt-in project-wide
// declaration-level diagnostic reported when every statically resolved call
// site of a Private/Friend Function or Property Get discards the result.
func (a Analyzer) functionAlwaysDiscardedFindings(ctx context.Context, files []parsedFile, projectViewComplete bool) ([]Finding, error) {
	// VBA258 claims every caller discards the result. Hidden source can
	// conceal a consuming caller or a dynamic-dispatch reference: modules
	// removed by PathFilter are invisible, and a parse error can swallow a
	// call anywhere (dynamic dispatch reaches Private procedures too). The
	// claim therefore requires a complete, cleanly parsed project view.
	if !projectViewComplete {
		return nil, nil
	}
	for i := range files {
		if files[i].IR.Parse.HasError || files[i].IR.Parse.HasMissing {
			return nil, nil
		}
	}
	byQualified := map[string]*alwaysDiscardedCandidate{}
	byName := map[string][]*alwaysDiscardedCandidate{}
	for fileIndex := range files {
		file := &files[fileIndex]
		interfaces := implementedInterfaceNames(file.IR)
		for procIndex := range file.IR.Procedures {
			proc := &file.IR.Procedures[procIndex]
			if !alwaysDiscardedEligible(proc, interfaces) {
				continue
			}
			candidate := &alwaysDiscardedCandidate{fileIndex: fileIndex, procIndex: procIndex}
			byQualified[strings.ToLower(proc.Symbol.QualifiedName)] = candidate
			name := strings.ToLower(cleanIdentifier(proc.Symbol.Name))
			byName[name] = append(byName[name], candidate)
		}
	}
	if len(byQualified) == 0 {
		return nil, nil
	}
	unknownDynamic := false
	for fileIndex := range files {
		if err := ctx.Err(); err != nil {
			return nil, err
		}
		file := &files[fileIndex]
		parse := symbols.ParseSummary{HasError: file.IR.Parse.HasError, HasMissing: file.IR.Parse.HasMissing}
		for procIndex := range file.IR.Procedures {
			proc := &file.IR.Procedures[procIndex]
			type calleeSpan struct {
				rng  vbaast.Range
				name string
			}
			calleeSpans := make([]calleeSpan, 0, len(proc.Calls))
			for _, call := range proc.Calls {
				name := strings.ToLower(call.Callee.Member)
				if name == "" {
					name = strings.ToLower(call.Callee.BaseName)
				}
				if tokenRange, ok := calleeTokenRange(call, file.Source); ok {
					calleeSpans = append(calleeSpans, calleeSpan{rng: tokenRange, name: name})
				}
				if call.IsRaiseEvent {
					continue
				}
				for _, ref := range calls.DynamicReferencesForIR(call, proc.Expressions, parse) {
					if strings.TrimSpace(ref.Target) == "" {
						// An opaque dynamic-dispatch target could reach any
						// candidate, so no all-discard claim survives.
						unknownDynamic = true
						continue
					}
					for _, candidate := range byName[dynamicReferenceName(ref.Target)] {
						candidate.uncertain = true
					}
				}
				if call.Resolution.Status == procedureir.ResolutionMatched && len(call.Resolution.Candidates) == 1 {
					// A uniquely resolved call is evidence only for its own
					// callee; calls resolved elsewhere prove nothing about
					// the candidates.
					if candidate, exists := byQualified[strings.ToLower(strings.TrimSpace(call.Resolution.Candidates[0].QualifiedName))]; exists {
						if callSiteDiscardsReturn(call, proc, file.Source) {
							candidate.discards++
						} else {
							candidate.consumed = true
						}
					}
					continue
				}
				// Member, ambiguous, unresolved, incomplete, external, and
				// dynamic calls may still reach a same-named candidate
				// (including With-block and late-bound receivers), so they
				// suppress rather than count.
				for _, candidate := range byName[strings.ToLower(call.Callee.BaseName)] {
					candidate.uncertain = true
				}
			}
			// Callee identifiers appear in Expressions but never in Accesses,
			// so an access is always a genuine value reference (argument,
			// receiver, or bare read) even inside a call-site range. For
			// expressions, only the callee's own name inside the callee token
			// range marks syntax rather than a reference; a same-named
			// argument such as the second Foo in "Foo Foo" stays a reference.
			isCalleePart := func(r vbaast.Range, name string) bool {
				for _, span := range calleeSpans {
					if r.StartByte >= span.rng.StartByte && r.EndByte <= span.rng.EndByte && name == span.name {
						return true
					}
				}
				return false
			}
			for _, access := range proc.Accesses {
				if access.Mode == procedureir.AccessWrite {
					continue
				}
				if access.Resolution.Status == procedureir.ResolutionMatched && len(access.Resolution.Candidates) == 1 {
					if candidate, exists := byQualified[strings.ToLower(strings.TrimSpace(access.Resolution.Candidates[0].QualifiedName))]; exists &&
						(candidate.fileIndex != fileIndex || candidate.procIndex != procIndex) {
						candidate.uncertain = true
					}
					continue
				}
				name := strings.ToLower(cleanIdentifier(discardedLastNamePart(access.Name)))
				candidates := byName[name]
				if len(candidates) == 0 {
					continue
				}
				if isCalleePart(access.Range, name) {
					continue
				}
				for _, candidate := range candidates {
					if candidate.fileIndex == fileIndex && candidate.procIndex == procIndex {
						continue
					}
					candidate.uncertain = true
				}
			}
			for _, expression := range proc.Expressions {
				if expression.Kind != procedureir.ExpressionIdentifier && expression.Kind != procedureir.ExpressionMember {
					continue
				}
				if expressionIsWriteTarget(proc, expression) {
					continue
				}
				name := strings.ToLower(cleanIdentifier(discardedLastNamePart(expression.Text)))
				candidates := byName[name]
				if len(candidates) == 0 {
					continue
				}
				if isCalleePart(expression.Range, name) {
					continue
				}
				for _, candidate := range candidates {
					if candidate.fileIndex == fileIndex && candidate.procIndex == procIndex {
						continue
					}
					candidate.uncertain = true
				}
			}
		}
	}
	if unknownDynamic {
		return nil, nil
	}
	type reportedCandidate struct {
		fileIndex int
		procIndex int
		discards  int
	}
	reported := make([]reportedCandidate, 0, len(byQualified))
	for _, candidate := range byQualified {
		if candidate.consumed || candidate.uncertain || candidate.discards == 0 {
			continue
		}
		reported = append(reported, reportedCandidate{fileIndex: candidate.fileIndex, procIndex: candidate.procIndex, discards: candidate.discards})
	}
	sort.Slice(reported, func(i, j int) bool {
		left, right := files[reported[i].fileIndex].Path, files[reported[j].fileIndex].Path
		if left != right {
			return left < right
		}
		return reported[i].procIndex < reported[j].procIndex
	})
	findings := make([]Finding, 0, len(reported))
	for _, entry := range reported {
		file := files[entry.fileIndex]
		procIR := &file.IR.Procedures[entry.procIndex]
		proc := sourceProcedure{
			Name:      procIR.Symbol.Name,
			StartLine: procIR.Symbol.DeclarationRange.StartLine,
			EndLine:   procIR.Symbol.DeclarationRange.EndLine,
		}
		if procedures := file.procedureView(); entry.procIndex < procedures.Len() {
			proc = procedures.valueAt(entry.procIndex)
		}
		kindLabel := "Function"
		if procIR.Symbol.Kind == procedureir.ProcedurePropertyGet {
			kindLabel = "Property Get"
		}
		finding := a.simpleFinding(
			file, proc, procIR.Symbol.DeclarationRange.StartLine, "VBA258", "information",
			fmt.Sprintf("Every known call site discards the return value of %s '%s' (%d call site(s)).", kindLabel, procIR.Symbol.QualifiedName, entry.discards),
			"All statically resolved callers invoke the procedure as a standalone statement and drop its result.",
			"Change the procedure to a Sub, or update at least one caller to consume the returned value.",
		)
		finding.ScopeEndLine = proc.EndLine
		findings = append(findings, finding)
	}
	return findings, ctx.Err()
}

// alwaysDiscardedEligible limits VBA258 to procedures whose callers are
// provably inside the project. Implicit visibility is Public in every module
// kind, so only explicit Private/Friend members qualify. Interface
// implementations (Interface_Member in a module declaring that interface with
// Implements) are invoked through the interface and are never referenced by
// name; unrelated underscored helpers in the same module keep coverage. The
// match uses the complete interface name because interface names may contain
// underscores: Implements I_Foo makes the Bar implementation I_Foo_Bar.
func alwaysDiscardedEligible(proc *procedureir.ProcedureIR, interfaces map[string]struct{}) bool {
	if proc == nil {
		return false
	}
	symbol := proc.Symbol
	if symbol.Kind != procedureir.ProcedureFunction && symbol.Kind != procedureir.ProcedurePropertyGet {
		return false
	}
	if symbol.Visibility != "Private" && symbol.Visibility != "Friend" {
		return false
	}
	if symbol.IsEventHandler || symbol.Recovered || len(symbol.ConditionalBranches) > 0 {
		return false
	}
	name := strings.ToLower(cleanIdentifier(symbol.Name))
	for iface := range interfaces {
		if strings.HasPrefix(name, iface+"_") {
			return false
		}
	}
	return true
}

// implementedInterfaceNames returns the lowered names a module declares with
// Implements statements.
func implementedInterfaceNames(document procedureir.DocumentIR) map[string]struct{} {
	names := map[string]struct{}{}
	for _, reference := range document.TypeReferences {
		if !strings.EqualFold(strings.TrimSpace(reference.Kind), "implements") {
			continue
		}
		if name := strings.ToLower(cleanIdentifier(discardedLastNamePart(reference.Target))); name != "" {
			names[name] = struct{}{}
		}
	}
	return names
}

// calleeTokenRange locates the callee's own source range inside the call's
// range. The call range also covers arguments, so a same-named argument like
// the second Foo in "Foo Foo" must remain a genuine reference; only the
// callee token itself is call syntax.
func calleeTokenRange(call procedureir.CallSite, source []byte) (vbaast.Range, bool) {
	text := call.Callee.Text
	if text == "" || call.Range.StartByte < 0 || call.Range.EndByte > len(source) || call.Range.StartByte > call.Range.EndByte {
		return vbaast.Range{}, false
	}
	offset := bytes.Index(source[call.Range.StartByte:call.Range.EndByte], []byte(text))
	if offset < 0 {
		return vbaast.Range{}, false
	}
	start := call.Range.StartByte + offset
	return vbaast.Range{StartByte: start, EndByte: start + len(text)}, true
}

func expressionIsWriteTarget(proc *procedureir.ProcedureIR, expression procedureir.Expression) bool {
	if expression.StatementID <= 0 || expression.StatementID > len(proc.Statements) {
		return false
	}
	return proc.Statements[expression.StatementID-1].TargetID == expression.ID
}

func discardedLastNamePart(text string) string {
	text = strings.TrimSpace(text)
	if index := strings.LastIndexAny(text, ".!"); index >= 0 {
		text = text[index+1:]
	}
	return text
}

func dynamicReferenceName(target string) string {
	name := discardedLastNamePart(target)
	name = strings.Trim(name, "\"'")
	return strings.ToLower(cleanIdentifier(name))
}
