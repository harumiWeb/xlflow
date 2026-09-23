package analyze

import (
	"fmt"
	"regexp"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/sourceproject"
)

// WorksheetCodeNameCatalog is an optional caller-owned capability that maps a
// worksheet's visible name to its VBA document-module CodeName.
type WorksheetCodeNameCatalog interface {
	Lookup(name string) (codeName string, ok bool)
}

type worksheetCodeNameFinding struct {
	Finding
	SheetName string
	CodeName  string
	Start     int
	End       int
}

var thisWorkbookWorksheetsLiteralRe = regexp.MustCompile(`(?i)\bThisWorkbook\s*\.\s*Worksheets\s*(?:\.\s*Item\s*)?\(`)

func sourceProjectHasWorksheetCodeNameCandidate(project sourceproject.SourceProject) bool {
	for _, file := range project.Files {
		lines := normalizedSourceLines(string(file.Source))
		for index := range lines {
			statement, start := worksheetLogicalStatement(lines, index, len(lines)-1)
			if start && worksheetCodeNameLiteralCandidate(statement) {
				return true
			}
		}
	}
	return false
}

func worksheetCodeNameLiteralCandidate(statement string) bool {
	masked := maskVBAStringsAndComments(statement)
	for _, match := range thisWorkbookWorksheetsLiteralRe.FindAllStringIndex(masked, -1) {
		open := match[1] - 1
		close := balancedCallEnd(statement, open)
		if close <= open {
			continue
		}
		if _, ok := worksheetNameLiteral(statement[open+1 : close]); ok {
			return true
		}
	}
	return false
}

func worksheetCodeNameFindings(statement string, catalog WorksheetCodeNameCatalog, documentModules map[string]string) []worksheetCodeNameFinding {
	if catalog == nil || len(documentModules) == 0 {
		return nil
	}
	masked := maskVBAStringsAndComments(statement)
	matches := thisWorkbookWorksheetsLiteralRe.FindAllStringIndex(masked, -1)
	findings := make([]worksheetCodeNameFinding, 0, len(matches))
	for _, match := range matches {
		open := match[1] - 1
		close := balancedCallEnd(statement, open)
		if close <= open {
			continue
		}
		sheetName, ok := worksheetNameLiteral(statement[open+1 : close])
		if !ok {
			continue
		}
		codeName, ok := catalog.Lookup(sheetName)
		if !ok || !worksheetCodeNameDocumentModuleExists(codeName, documentModules) {
			continue
		}
		findings = append(findings, worksheetCodeNameFinding{
			Finding: Finding{
				Code:       "VBA260",
				Severity:   "warning",
				Message:    fmt.Sprintf("ThisWorkbook.Worksheets(%q) uses a visible sheet name instead of the stable CodeName %s.", sheetName, codeName),
				Reason:     "A worksheet CodeName remains stable when the visible sheet tab is renamed.",
				Suggestion: fmt.Sprintf("Use %s directly for this worksheet.", codeName),
			},
			SheetName: sheetName,
			CodeName:  codeName,
			Start:     match[0],
			End:       close + 1,
		})
	}
	return findings
}

func (a Analyzer) worksheetCodeNameStatementFindings(file parsedFile, proc sourceProcedure, line int, statement string, documentModules map[string]string) []Finding {
	candidates := worksheetCodeNameFindings(statement, a.WorksheetCodeNames, documentModules)
	out := make([]Finding, 0, len(candidates))
	for _, candidate := range candidates {
		finding := a.simpleFinding(file, proc, line, candidate.Code, candidate.Severity, candidate.Message, candidate.Reason, candidate.Suggestion)
		finding.Column = candidate.Start + 1
		finding.EndLine = line
		finding.EndColumn = candidate.End + 1
		out = append(out, finding)
	}
	return out
}

func worksheetCodeNameDocumentModuleExists(codeName string, documentModules map[string]string) bool {
	found := false
	for key, value := range documentModules {
		candidate := strings.TrimSpace(value)
		if candidate == "" {
			candidate = strings.TrimSpace(key)
		}
		if !strings.EqualFold(candidate, codeName) {
			continue
		}
		if found {
			return false
		}
		found = true
	}
	return found
}

func maskVBAStringsAndComments(text string) string {
	masked := []byte(text)
	inString := false
	atStatementStart := true
	for index := 0; index < len(masked); index++ {
		if inString {
			if masked[index] == '"' {
				if index+1 < len(masked) && masked[index+1] == '"' {
					masked[index], masked[index+1] = ' ', ' '
					index++
					continue
				}
				masked[index] = ' '
				inString = false
				continue
			}
			masked[index] = ' '
			continue
		}
		if masked[index] == '"' {
			masked[index] = ' '
			inString = true
			continue
		}
		if atStatementStart && remCommentStart(masked, index) {
			for commentIndex := index; commentIndex < len(masked); commentIndex++ {
				masked[commentIndex] = ' '
			}
			break
		}
		if masked[index] == '\'' {
			for commentIndex := index; commentIndex < len(masked); commentIndex++ {
				masked[commentIndex] = ' '
			}
			break
		}
		if masked[index] == ':' {
			atStatementStart = true
			continue
		}
		if masked[index] != ' ' && masked[index] != '\t' {
			atStatementStart = false
		}
	}
	return string(masked)
}

func remCommentStart(text []byte, index int) bool {
	if index+3 > len(text) || !strings.EqualFold(string(text[index:index+3]), "rem") {
		return false
	}
	return index+3 == len(text) || text[index+3] == ' ' || text[index+3] == '\t'
}
