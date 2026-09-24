package analyze

import (
	"regexp"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// unusedParameterFindings reports parameters of private procedures that are
// never referenced in the procedure body. Only Private procedures are
// candidates: Public and Friend members are callable across module or project
// boundaries, so their signatures are fixed by callers the analyzer cannot
// enumerate. Implements members, event handlers, and procedures reachable
// through string-based dynamic invocation are excluded for the same reason.
// dynamicNames is the project-wide set of identifier tokens found inside
// string literals (Application.Run "Module.Proc", OnTime, OnAction,
// CallByName). It is built once per analysis run from every visible module;
// a nil set falls back to the candidate file's own literals.
func (a Analyzer) unusedParameterFindings(file parsedFile, proc sourceProcedure, dynamicNames map[string]bool) []Finding {
	if !unusedParameterEligibleProcedure(file, proc, dynamicNames) {
		return nil
	}
	var findings []Finding
	if proc.IR == nil {
		return nil
	}
	matchers := make(map[string]*regexp.Regexp)
	for _, parameter := range proc.IR.Symbol.Parameters {
		name := cleanIdentifier(parameter.Name)
		if name == "" || isIgnoredUnusedName(name) {
			continue
		}
		matcher := matchers[name]
		if matcher == nil {
			matcher = unusedNameMatcher(name)
			matchers[name] = matcher
		}
		if parameterIdentifierUsed(file, proc, parameter, matcher) {
			continue
		}
		line := parameter.Range.StartLine
		if line < 1 {
			line = proc.IR.Symbol.DeclarationRange.StartLine
		}
		finding := a.simpleFinding(file, proc, line, "VBA260", "warning",
			"Parameter "+name+" is never used inside the procedure body.",
			"No statement reads or writes this parameter; callers pass a value the implementation ignores.",
			"Remove the parameter and update call sites, or use it where the signature requires it.")
		finding.Column = parameter.Range.StartColumn
		finding.EndLine = parameter.Range.EndLine
		finding.EndColumn = parameter.Range.EndColumn
		findings = append(findings, finding)
	}
	return findings
}

func unusedParameterEligibleProcedure(file parsedFile, proc sourceProcedure, dynamicNames map[string]bool) bool {
	if proc.IR == nil {
		return false
	}
	symbol := proc.IR.Symbol
	if symbol.Recovered || len(symbol.ConditionalBranches) > 0 {
		return false
	}
	if !strings.EqualFold(strings.TrimSpace(symbol.Visibility), "private") {
		return false
	}
	if symbol.IsEventHandler || eventHandlerKind(file, proc) != "" {
		return false
	}
	facts := file.moduleAnalysisFacts().unusedDeclarationFacts(file.Lines)
	if signatureConstrainedEvent(file, proc, facts) {
		return false
	}
	name := strings.ToLower(cleanIdentifier(symbol.Name))
	for _, iface := range facts.implementsTargets {
		if strings.HasPrefix(name, iface+"_") {
			return false
		}
	}
	// A string literal carrying this procedure's name marks a discoverable
	// dynamic entry point (Application.Run, OnTime, OnAction, CallByName).
	// The project-wide set subsumes this file's literals; a nil set means the
	// caller had no project view and falls back to file-local literals.
	if dynamicNames == nil {
		dynamicNames = facts.literalNames
	}
	if dynamicNames[name] || dynamicNames[strings.ToLower(file.Module+"."+cleanIdentifier(symbol.Name))] {
		return false
	}
	return true
}

// signatureConstrainedEvent reports whether the procedure name follows an
// event-binding shape whose signature is fixed by a host the analyzer cannot
// enumerate. WithEvents callbacks bind as <field>_<event> in class, form, and
// document modules; document modules additionally host control and host-object
// events (worksheet ActiveX controls, Access form/report sections) that are not
// statically discoverable, so any <object>_<event> name there is treated as a
// constrained signature. Both directions only suppress findings.
func signatureConstrainedEvent(file parsedFile, proc sourceProcedure, facts *unusedDeclFileFacts) bool {
	name := strings.ToLower(cleanIdentifier(proc.IR.Symbol.Name))
	index := strings.LastIndex(name, "_")
	if index <= 0 || index == len(name)-1 {
		return false
	}
	if !facts.withEventsComplete {
		// An unparseable WithEvents declaration means any underscored name in
		// this module could be a field callback.
		return true
	}
	for field := range facts.withEventsFields {
		// A WithEvents field may itself begin with Test (for example
		// TestApp_WindowBeforeDoubleClick), so field matching must run before
		// the test-name exemption below.
		if strings.HasPrefix(name, field+"_") {
			return true
		}
	}
	if strings.HasPrefix(name, "test") || strings.HasSuffix(name, "_test") {
		return false
	}
	return strings.EqualFold(file.ModuleKind, "document")
}

// parameterIdentifierUsed scans the procedure text for a whole-word
// occurrence of the parameter name outside the parameter's own declarator
// range. A lexical scan is intentionally permissive: occurrences inside
// comments or string literals also mark the parameter used, which only ever
// suppresses a finding.
func parameterIdentifierUsed(file parsedFile, proc sourceProcedure, parameter procedureir.Parameter, matcher *regexp.Regexp) bool {
	start := proc.StartLine - 1
	if start < 0 {
		start = 0
	}
	end := proc.EndLine
	if end > len(file.Lines) {
		end = len(file.Lines)
	}
	for i := start; i < end; i++ {
		line := file.Lines[i]
		for _, loc := range matcher.FindAllStringIndex(line, -1) {
			begin := loc[0]
			// The match may include one delimiter character on each side.
			if !isIdentifierByte(line[begin]) {
				begin++
			}
			if positionInRange(i+1, begin+1, parameter.Range) {
				continue
			}
			return true
		}
	}
	return false
}

// positionInRange reports whether a 1-based line and 1-based column fall
// inside a source range. Used to skip a name's own declarator when scanning
// lines for textual uses.
func positionInRange(line, column int, rng vbaast.Range) bool {
	if line < rng.StartLine || line > rng.EndLine {
		return false
	}
	if line == rng.StartLine && column < rng.StartColumn {
		return false
	}
	if line == rng.EndLine && column >= rng.EndColumn {
		return false
	}
	return true
}

func isIdentifierByte(b byte) bool {
	return b == '_' || b >= '0' && b <= '9' || b >= 'a' && b <= 'z' || b >= 'A' && b <= 'Z' || b >= 0x80
}

func unusedNameMatcher(name string) *regexp.Regexp {
	return regexp.MustCompile(`(?i)(^|[^0-9A-Za-z_])` + regexp.QuoteMeta(name) + `([^0-9A-Za-z_]|$)`)
}

// moduleImplementsTargets returns the lowercased interface names declared by
// Implements statements. A qualified target (`Implements Lib.IFace`) binds
// members as <interface>_<member> using the final segment only, so the
// qualifier is stripped. Implements members must keep their signature even
// when the implementation is Private.
func moduleImplementsTargets(lines []string) []string {
	var targets []string
	for _, line := range lines {
		fields := strings.Fields(strings.ToLower(normalizedCodeLine(line)))
		if len(fields) == 2 && fields[0] == "implements" {
			target := strings.Trim(fields[1], "[]")
			if dot := strings.LastIndexByte(target, '.'); dot >= 0 {
				target = strings.Trim(target[dot+1:], "[]")
			}
			targets = append(targets, target)
		}
	}
	return targets
}

func stringLiteralContainsName(literal, name string) bool {
	for i := 0; i+len(name) <= len(literal); i++ {
		if !strings.EqualFold(literal[i:i+len(name)], name) {
			continue
		}
		if i > 0 && isIdentifierByte(literal[i-1]) {
			continue
		}
		end := i + len(name)
		if end < len(literal) && isIdentifierByte(literal[end]) {
			continue
		}
		return true
	}
	return false
}

func isIgnoredUnusedName(name string) bool {
	name = strings.ToLower(strings.TrimSpace(name))
	return name == "" || name == "_" || strings.HasPrefix(name, "unused") || strings.HasPrefix(name, "ignore")
}
