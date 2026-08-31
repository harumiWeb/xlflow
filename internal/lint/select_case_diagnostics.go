package lint

import (
	"strings"

	"github.com/harumiWeb/xlflow/internal/gui"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

// selectCaseDiagnosticCode identifies the Select/Case syntax family.  The
// registry metadata remains separate from this parser helper.
const selectCaseDiagnosticCode = "VB063"

// selectCaseSyntaxIssue is the tree-independent payload emitted by
// selectCaseSyntaxIssues.  The linter integration can convert Range and the
// message fields into an Issue without retaining a tree-sitter node.
//
// The helper intentionally returns a small payload rather than Issue: this
// keeps the syntax interpretation testable before it is connected to the
// linter's suppression and precedence pipeline.
type selectCaseSyntaxIssue struct {
	Code       string
	Kind       string
	Range      vbaast.Range
	Message    string
	Suggestion string
}

// selectCaseSyntaxIssues identifies only high-confidence Select/Case ordering
// errors.  It recognizes statement boundaries lexically (including colon
// separated statements), while using the parser root as a recovery gate.  If
// the tree contains any error or missing node, the construct is ambiguous and
// the generic parser-recovery diagnostic remains the sole owner.
//
// The returned issues are in source order.  Their ranges cover the offending
// colon-separated statement, which gives a future linter adapter a stable
// location without retaining a tree-sitter node.
func selectCaseSyntaxIssues(source string, root *tree_sitter.Node) []selectCaseSyntaxIssue {
	if root == nil || root.HasError() || vbaast.HasMissing(root) {
		return nil
	}

	// Keep the original physical-line bytes so ranges remain offsets into the
	// caller's source (not a CRLF-normalized copy).
	lines := strings.Split(source, "\n")
	frames := make([]selectCaseFrame, 0, 4)
	var issues []selectCaseSyntaxIssue
	lineOffset := 0

	for lineIndex, line := range lines {
		// Strip comments before splitting on colons.  gui.StripComment is aware
		// of doubled quotes, so words in string literals cannot become control
		// statements accidentally.
		parts := splitStatementsWithColumns(gui.StripComment(line))
		for _, part := range parts {
			text, start := normalizeSelectCasePart(part)
			if text == "" {
				continue
			}
			first, rest := firstVBAWord(text)
			// Rem starts a comment when it is the first token of a
			// colon-separated statement; the rest of the physical line is
			// comment text rather than additional VBA statements.
			if strings.EqualFold(first, "Rem") {
				break
			}
			switch strings.ToLower(first) {
			case "select":
				if selectCaseHeader(rest) {
					frames = append(frames, selectCaseFrame{})
				}
			case "end":
				if selectCaseCloser(rest) && len(frames) > 0 {
					frames = frames[:len(frames)-1]
				}
			case "case":
				issue, ok := selectCaseBranchIssue(frames, rest, lineIndex+1, start+1, len(text), lineOffset+start, lineOffset+start+len(text))
				if ok {
					issues = append(issues, issue)
				}
				// A Case branch is part of the current Select frame even when
				// it is invalid.  Keeping that state lets us report one precise
				// issue per subsequent branch rather than cascading outer-frame
				// guesses.
				if len(frames) > 0 {
					if caseElse(rest) {
						frames[len(frames)-1].sawElse = true
					}
				}
			}
		}
		lineOffset += len(line) + 1 // include the split '\n' when present
	}
	return issues
}

type selectCaseFrame struct {
	sawElse bool
}

// normalizeSelectCasePart removes an optional numeric VBA line label and
// returns the statement text with its byte offset within the physical line.
// splitStatementsWithColumns has already trimmed the outer whitespace, but a
// numeric label may still precede the first keyword (for example `100 Case 1`).
func normalizeSelectCasePart(part statementPart) (string, int) {
	text := part.text
	start := part.start
	if n := leadingLineNumberLength(text); n > 0 {
		text = text[n:]
		start += n
	}
	leading := len(text) - len(strings.TrimLeft(text, " \t"))
	start += leading
	text = strings.TrimSpace(text)
	return text, start
}

func selectCaseHeader(rest string) bool {
	word, _ := firstVBAWord(rest)
	if !strings.EqualFold(word, "Case") {
		return false
	}
	// `Select Case` is a complete header even when the expression is
	// continued on the next physical line.  The parser recovery gate above
	// handles genuinely malformed or incomplete headers.
	return true
}

func selectCaseCloser(rest string) bool {
	word, remainder := firstVBAWord(rest)
	return strings.EqualFold(word, "Select") && strings.TrimSpace(remainder) == ""
}

func caseElse(rest string) bool {
	word, remainder := firstVBAWord(rest)
	return strings.EqualFold(word, "Else") && strings.TrimSpace(remainder) == ""
}

func selectCaseBranchIssue(frames []selectCaseFrame, rest string, line, column, textLength, start, end int) (selectCaseSyntaxIssue, bool) {
	range_ := vbaast.Range{
		StartLine: line, StartColumn: column,
		EndLine: line, EndColumn: column + textLength,
		StartByte: start, EndByte: end,
	}
	if len(frames) == 0 {
		return selectCaseSyntaxIssue{
			Code:       selectCaseDiagnosticCode,
			Kind:       "case_outside_select",
			Range:      range_,
			Message:    "A Case statement must appear inside a Select Case block.",
			Suggestion: "Move this Case into a Select Case block.",
		}, true
	}
	frame := frames[len(frames)-1]
	if caseElse(rest) {
		if frame.sawElse {
			return selectCaseSyntaxIssue{
				Code:       selectCaseDiagnosticCode,
				Kind:       "duplicate_case_else",
				Range:      range_,
				Message:    "A Select Case block may contain only one Case Else branch.",
				Suggestion: "Remove the duplicate Case Else branch.",
			}, true
		}
		return selectCaseSyntaxIssue{}, false
	}
	if frame.sawElse {
		return selectCaseSyntaxIssue{
			Code:       selectCaseDiagnosticCode,
			Kind:       "case_after_else",
			Range:      range_,
			Message:    "No Case branch may follow Case Else.",
			Suggestion: "Move this Case before Case Else.",
		}, true
	}
	return selectCaseSyntaxIssue{}, false
}
