package analyze

import (
	"regexp"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/intel"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

var (
	excelAPIContractAssignmentRe  = regexp.MustCompile(`(?i)^\s*(?:let\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=`)
	excelAPIContractFunctionRe    = regexp.MustCompile(`(?i)^\s*(?:public|private|friend)?\s*function\s+([A-Za-z_][A-Za-z0-9_]*)\b.*\bas\s+boolean\b`)
	excelAPIContractAnyFunctionRe = regexp.MustCompile(`(?i)^\s*(?:(public|private|friend)\s+)?function\s+([A-Za-z_][A-Za-z0-9_]*)\b`)
)

// excelAPIFailureContractFindings adapts the typed, member-resolved call
// diagnostics to analyzer findings and applies the procedure-local safeguards
// that need source control-flow context.
func (a Analyzer) excelAPIFailureContractFindings(file parsedFile) []Finding {
	if !a.Config.Analyze.DetectExcelAPIFailureContracts || a.typeDB == nil {
		return nil
	}
	diagnostics := (intel.Analyzer{RootDir: a.RootDir, Config: a.Config, DB: a.typeDB}).ExcelAPIFailureContractDiagnostics(intel.Document{
		Path: file.Path, Source: string(file.Source),
	})
	if len(diagnostics) == 0 {
		return nil
	}
	procedures := sourceProceduresFromIR(file.IR, file.CFG)
	aliases := a.errorGuardAliases
	if aliases == nil {
		aliases = isErrorGuardAliases(file.Lines)
	}
	out := make([]Finding, 0, len(diagnostics))
	for _, diagnostic := range diagnostics {
		line := diagnostic.Range.Start.Line + 1
		proc := procedureForLine(procedures, line, len(file.Lines))
		if isExceptionContractDiagnostic(diagnostic.Message) && exceptionContractHasErrorStrategy(file.Lines, proc, line) {
			continue
		}
		if isVariantErrorContractDiagnostic(diagnostic.Message) && variantContractResultSafelyConsumed(file.Lines, proc, line, aliases) {
			continue
		}
		finding := a.simpleFinding(
			file, proc, line, diagnostic.Code, diagnostic.Severity, diagnostic.Message,
			excelAPIContractReason(diagnostic.Message), excelAPIContractSuggestion(diagnostic.Message),
		)
		finding.Column = diagnostic.Range.Start.Character + 1
		finding.EndLine = diagnostic.Range.End.Line + 1
		finding.EndColumn = diagnostic.Range.End.Character + 1
		out = append(out, finding)
	}
	return out
}

func procedureForLine(procedures []sourceProcedure, line, lineCount int) sourceProcedure {
	for _, procedure := range procedures {
		if line >= procedure.StartLine && line <= procedure.EndLine {
			return procedure
		}
	}
	return sourceProcedure{StartLine: 1, EndLine: lineCount}
}

func isExceptionContractDiagnostic(message string) bool {
	return strings.HasPrefix(message, "Range.SpecialCells may raise") || strings.HasPrefix(message, "WorksheetFunction.")
}

func isVariantErrorContractDiagnostic(message string) bool {
	return strings.HasPrefix(message, "Application.") && strings.Contains(message, "Variant/Error")
}

func excelAPIContractReason(message string) string {
	if isExceptionContractDiagnostic(message) {
		return "This Excel API can raise a runtime error when its requested result does not exist; the call is not proof that an error will occur."
	}
	return "This Excel API can return a Variant/Error when no result exists; consuming that value as an ordinary result can raise a later runtime error."
}

func excelAPIContractSuggestion(message string) string {
	if isExceptionContractDiagnostic(message) {
		return "Use a local On Error handler, or keep On Error Resume Next to this probe, inspect Err, and restore error handling immediately."
	}
	return "Store the result and use IsError before consuming its success value."
}

func exceptionContractHasErrorStrategy(lines []string, proc sourceProcedure, line int) bool {
	mode := ""
	for index := proc.StartLine - 1; index < line && index < len(lines); index++ {
		statement := strings.ToLower(strings.TrimSpace(normalizedCodeLine(lines[index])))
		switch {
		case strings.HasPrefix(statement, "on error resume next"):
			mode = "resume-next"
		case strings.HasPrefix(statement, "on error goto 0"):
			mode = ""
		case strings.HasPrefix(statement, "on error goto "):
			mode = "handler"
		}
	}
	if mode == "handler" {
		return true
	}
	if mode != "resume-next" {
		return false
	}
	return narrowResumeNextProbe(lines, proc, line)
}

func narrowResumeNextProbe(lines []string, proc sourceProcedure, line int) bool {
	checkedErr := false
	steps := 0
	for index := line; index < proc.EndLine && index < len(lines); index++ {
		statement := strings.ToLower(strings.TrimSpace(normalizedCodeLine(lines[index])))
		if statement == "" {
			continue
		}
		steps++
		if strings.Contains(statement, "err.number") {
			checkedErr = true
		}
		if strings.HasPrefix(statement, "on error goto 0") {
			return checkedErr
		}
		if strings.HasPrefix(statement, "on error ") || steps > 4 {
			return false
		}
	}
	return false
}

func isErrorGuardAliases(lines []string) map[string]bool {
	aliases := map[string]bool{}
	function := ""
	for _, line := range lines {
		statement := strings.TrimSpace(normalizedCodeLine(line))
		if match := excelAPIContractFunctionRe.FindStringSubmatch(statement); len(match) > 0 {
			function = strings.ToLower(match[1])
			continue
		}
		if strings.HasPrefix(strings.ToLower(statement), "end function") {
			function = ""
			continue
		}
		if function != "" && strings.Contains(strings.ToLower(statement), function+" = iserror(") {
			aliases[function] = true
		}
	}
	return aliases
}

// projectIsErrorGuardAliases returns only uniquely named project-local Boolean
// wrappers whose body delegates to IsError. Batch analysis can safely use those
// aliases across modules; real-time analysis intentionally supplies no project
// summary and therefore remains document-local.
func projectIsErrorGuardAliases(files []parsedFile) map[string]bool {
	candidates := map[string]bool{}
	counts := map[string]int{}
	for _, file := range files {
		for name := range isErrorGuardAliases(file.Lines) {
			candidates[name] = true
			counts[name]++
		}
	}
	aliases := map[string]bool{}
	for name := range candidates {
		if counts[name] == 1 {
			aliases[name] = true
		}
	}
	return aliases
}

// errorValueWrapperFindings treats calls to a local Function that catches an
// exceptional Excel API failure and returns CVErr as a Variant/Error producer.
// The wrapper's own exception is handled, but an unchecked returned sentinel
// remains unsafe at the caller.
func (a Analyzer) errorValueWrapperFindings(file parsedFile) []Finding {
	if !a.Config.Analyze.DetectExcelAPIFailureContracts {
		return nil
	}
	wrappers := a.errorValueWrappers
	if wrappers == nil {
		wrappers = errorValueWrappersForFile(file, false)
	}
	if len(wrappers) == 0 {
		return nil
	}
	procedures := sourceProceduresFromIR(file.IR, file.CFG)
	aliases := a.errorGuardAliases
	if aliases == nil {
		aliases = isErrorGuardAliases(file.Lines)
	}
	var findings []Finding
	for _, irProcedure := range file.IR.Procedures {
		proc := procedureForLine(procedures, irProcedure.Symbol.DeclarationRange.StartLine+1, len(file.Lines))
		for _, call := range irProcedure.Calls {
			if !wrapperCallMatches(call, wrappers) {
				continue
			}
			line := call.Range.StartLine
			if wrapperCallIsDirectIsError(file, call) || discardedWrapperCallResult(file, call) || variantContractResultSafelyConsumed(file.Lines, proc, line, aliases) {
				continue
			}
			finding := a.simpleFinding(
				file, proc, line, "VBA218", "warning",
				"Local wrapper "+call.Callee.BaseName+" may return a Variant/Error sentinel; check it with IsError before consuming the value.",
				"The wrapper handles its internal Excel exception by returning CVErr, so its caller can still receive a Variant/Error instead of an ordinary result.",
				"Store the wrapper result and use IsError before consuming its success value.",
			)
			finding.Column = call.Range.StartColumn
			finding.EndLine = call.Range.EndLine
			finding.EndColumn = call.Range.EndColumn
			findings = append(findings, finding)
		}
	}
	return findings
}

func projectErrorValueWrappers(files []parsedFile) map[string]bool {
	wrappers := map[string]bool{}
	for _, file := range files {
		for key := range errorValueWrappersForFile(file, false) {
			wrappers[key] = true
		}
	}
	return wrappers
}

func errorValueWrappersForFile(file parsedFile, publicOnly bool) map[string]bool {
	wrappers := map[string]bool{}
	name := ""
	visibility := ""
	hasHandler := false
	returnsError := false
	for _, line := range file.Lines {
		statement := strings.TrimSpace(normalizedCodeLine(line))
		if match := excelAPIContractAnyFunctionRe.FindStringSubmatch(statement); len(match) > 0 {
			name, visibility = match[2], strings.ToLower(match[1])
			hasHandler, returnsError = false, false
			continue
		}
		if name == "" {
			continue
		}
		lower := strings.ToLower(statement)
		if strings.HasPrefix(lower, "on error goto ") && !strings.HasPrefix(lower, "on error goto 0") {
			hasHandler = true
		}
		if strings.Contains(lower, strings.ToLower(name)+" = cverr(") {
			returnsError = true
		}
		if strings.HasPrefix(lower, "end function") {
			if hasHandler && returnsError && (!publicOnly || visibility == "public") {
				wrappers[strings.ToLower(file.Module+"."+name)] = true
			}
			name = ""
		}
	}
	return wrappers
}

func wrapperCallMatches(call procedureir.CallSite, wrappers map[string]bool) bool {
	if call.Callee.Receiver != nil {
		return false
	}
	return call.Resolution.Status == procedureir.ResolutionMatched &&
		len(call.Resolution.Candidates) == 1 &&
		wrappers[strings.ToLower(call.Resolution.Candidates[0].QualifiedName)]
}

func wrapperCallIsDirectIsError(file parsedFile, call procedureir.CallSite) bool {
	line := sourceLine(file.Lines, call.Range.StartLine-1)
	pattern := regexp.MustCompile(`(?i)\biserror\s*\(\s*` + regexp.QuoteMeta(call.Callee.BaseName) + `\s*(?:\(|\b)`)
	return pattern.MatchString(line)
}

func discardedWrapperCallResult(file parsedFile, call procedureir.CallSite) bool {
	line := strings.TrimSpace(strings.ToLower(sourceLine(file.Lines, call.Range.StartLine-1)))
	line = strings.TrimSpace(strings.TrimPrefix(line, "call "))
	base := strings.ToLower(call.Callee.BaseName)
	return (strings.HasPrefix(line, base+"(") || strings.HasPrefix(line, base+" ")) && !strings.Contains(line, "=")
}

func sourceLine(lines []string, zeroBasedLine int) string {
	if zeroBasedLine < 0 || zeroBasedLine >= len(lines) {
		return ""
	}
	return normalizedCodeLine(lines[zeroBasedLine])
}

func variantContractResultSafelyConsumed(lines []string, proc sourceProcedure, line int, aliases map[string]bool) bool {
	if line < 1 || line > len(lines) {
		return false
	}
	match := excelAPIContractAssignmentRe.FindStringSubmatch(normalizedCodeLine(lines[line-1]))
	if len(match) == 0 {
		return false
	}
	name := match[1]
	nameRe := regexp.MustCompile(`(?i)\b` + regexp.QuoteMeta(name) + `\b`)
	guardRe := regexp.MustCompile(`(?i)^\s*if\s+(not\s+)?(?:` + strings.Join(sortedKeys(aliases), "|") + `)\s*\(\s*` + regexp.QuoteMeta(name) + `\s*\)\s+then\b`)
	// A plain assignment is harmless until it is consumed. Suppress its call-site
	// diagnostic when every observed use is inside a known success path.
	safeDepth := 0
	errorGuardDepth := 0
	postErrorExit := false
	for index := line; index < proc.EndLine && index < len(lines); index++ {
		statement := normalizedCodeLine(lines[index])
		lower := strings.ToLower(strings.TrimSpace(statement))
		if excelAPIContractAssignmentRe.MatchString(statement) && strings.EqualFold(excelAPIContractAssignmentRe.FindStringSubmatch(statement)[1], name) {
			break
		}
		guardMatch := guardRe.FindStringSubmatch(statement)
		guarded := len(guardMatch) > 0
		if guarded {
			if strings.TrimSpace(guardMatch[1]) != "" {
				safeDepth++
			} else {
				errorGuardDepth++
				if strings.Contains(lower, "exit sub") || strings.Contains(lower, "exit function") || strings.Contains(lower, "exit property") {
					postErrorExit = true
				}
			}
			continue
		}
		if strings.HasPrefix(lower, "else") && errorGuardDepth > 0 {
			safeDepth++
			continue
		}
		if errorGuardDepth > 0 && safeDepth == 0 && (lower == "exit sub" || lower == "exit function" || lower == "exit property") {
			postErrorExit = true
			continue
		}
		if strings.HasPrefix(lower, "end if") {
			if safeDepth > 0 {
				safeDepth--
			}
			if errorGuardDepth > 0 {
				errorGuardDepth--
			}
			continue
		}
		if !nameRe.MatchString(statement) || guarded {
			continue
		}
		if safeDepth == 0 && !postErrorExit {
			return false
		}
	}
	return true
}

func sortedKeys(values map[string]bool) []string {
	if len(values) == 0 {
		return []string{"iserror"}
	}
	keys := make([]string, 0, len(values)+1)
	keys = append(keys, "iserror")
	for key := range values {
		keys = append(keys, key)
	}
	return keys
}
