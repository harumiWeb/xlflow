package lint

import (
	"bytes"
	"strings"
	"unicode"

	"github.com/harumiWeb/xlflow/internal/gui"
	"github.com/harumiWeb/xlflow/internal/vba/ast"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

const procedureTerminatorStyleDiagnosticCode = "VB066"

// DefaultProcedureBoundaryClassifier is the compatibility table promoted from
// the local VBE audit in testdata/vbe-oracle.  The table deliberately keys on
// the normalized opener/accessor/terminator rather than tree-sitter node
// shape: parser structure, VBA compiler validity, and xlflow style preference
// are separate contracts.
//
// Excel 16.0 (build 17932, x64, ja-JP) accepted Property Get closed by End
// Sub or End Function in all audited module hosts.  It rejected every other
// mismatched combination in the 5x3 opener/terminator matrix.  Matching
// forms never reach this classifier.
func DefaultProcedureBoundaryClassifier(context ProcedureBoundaryContext) ProcedureBoundaryDecision {
	// UserForm modules are intentionally outside issue #597's evidence set.
	// Do not generalize the document-module result to an unaudited host.
	switch strings.ToLower(strings.TrimSpace(context.ModuleKind)) {
	case "", "standard", "class", "document", "document-workbook", "document-worksheet":
	default:
		return ProcedureBoundaryDecision{Outcome: ProcedureBoundaryCompileInvalid}
	}
	if strings.EqualFold(context.OpenerKind, "Property") &&
		strings.EqualFold(context.Accessor, "Get") &&
		(strings.EqualFold(context.TerminatorKind, "Sub") || strings.EqualFold(context.TerminatorKind, "Function")) {
		return ProcedureBoundaryDecision{
			Outcome:  ProcedureBoundaryStyleMismatch,
			Code:     procedureTerminatorStyleDiagnosticCode,
			Severity: "warning",
			Message:  "VBE accepts this Property Get terminator, but End Property is the canonical style for the accessor.",
		}
	}
	return ProcedureBoundaryDecision{Outcome: ProcedureBoundaryCompileInvalid}
}

// IsAcceptedProcedureBoundaryRecovery reports the one parser-recovery shape
// covered by the promoted VBE evidence: a Property Get closed by End Sub or
// End Function in an audited module kind. It is intentionally conservative;
// the parse tree must contain exactly one recovery range, an error node for
// the exact accepted closer token, and no independent recovery node or code
// after that token on the same line. Other parser errors remain hard parse
// failures.
// The helper is used by the batch analyzer's parse gate so VBE-accepted source
// can reach the normal non-blocking VB066 lint projection.
func IsAcceptedProcedureBoundaryRecovery(root *tree_sitter.Node, source []byte, moduleKind string) bool {
	if root == nil {
		return false
	}
	problemRanges := parseProblemRanges(root)
	if len(problemRanges) != 1 {
		return false
	}

	var accepted []ProcedureBoundaryContext
	classifier := func(context ProcedureBoundaryContext) ProcedureBoundaryDecision {
		decision := DefaultProcedureBoundaryClassifier(context)
		if decision.Outcome == ProcedureBoundaryStyleMismatch {
			accepted = append(accepted, context)
			// The tracker is only being used to identify the source boundary;
			// the parser recovery decision is made against the exact range below.
			return ProcedureBoundaryDecision{Outcome: ProcedureBoundaryAccepted}
		}
		return decision
	}
	tracker := newProcedureBoundaryTracker(Linter{ModuleKind: moduleKind, ProcedureBoundaryClassifier: classifier}, "<parser-recovery>")
	lines := strings.Split(strings.ReplaceAll(string(source), "\r\n", "\n"), "\n")
	for lineIndex, line := range lines {
		lineNo := lineIndex + 1
		code := gui.StripComment(line)
		trimmed := strings.TrimSpace(code)
		lower := strings.ToLower(trimmed)
		if isConditionalCompilationDirective(lower) {
			tracker.processDirective(lineNo, lower)
			continue
		}
		tracker.processStatement(lineNo, code)
	}
	if tracker.ambiguous || len(tracker.conditionals) != 0 || len(accepted) != 1 {
		return false
	}
	problem := problemRanges[0]
	return acceptedProcedureBoundaryRecoveryMatches(root, source, problem, accepted[0])
}

func acceptedProcedureBoundaryRecoveryMatches(root *tree_sitter.Node, source []byte, problem ast.Range, context ProcedureBoundaryContext) bool {
	line, lineStart, ok := sourceLineAt(source, context.ClosingLine)
	if !ok {
		return false
	}
	code := gui.StripComment(string(line))
	trimmed := strings.TrimSpace(code)
	fields := strings.Fields(trimmed)
	if len(fields) != 2 || !strings.EqualFold(fields[0], "End") || !strings.EqualFold(fields[1], context.TerminatorKind) {
		// A colon-separated statement, even one tree-sitter can recover as a
		// legal-looking node, is an independent construct and must not be
		// admitted by the single-boundary parser exception.
		return false
	}
	leading := len(code) - len(strings.TrimLeftFunc(code, unicode.IsSpace))
	tokenStart := lineStart + leading
	tokenEnd := tokenStart + len(trimmed)
	if problem.StartByte > tokenStart || problem.EndByte < tokenEnd {
		return false
	}
	return hasOnlyAcceptedProcedureTerminatorRecovery(root, problem, tokenStart, tokenEnd)
}

func sourceLineAt(source []byte, lineNo int) ([]byte, int, bool) {
	if lineNo < 1 {
		return nil, 0, false
	}
	start := 0
	for current := 1; start <= len(source); current++ {
		end := bytes.IndexByte(source[start:], '\n')
		hasNewline := end >= 0
		if !hasNewline {
			end = len(source) - start
		}
		lineEnd := start + end
		line := bytes.TrimSuffix(source[start:lineEnd], []byte{'\r'})
		if current == lineNo {
			return line, start, true
		}
		if !hasNewline {
			return nil, 0, false
		}
		start = lineEnd + 1
	}
	return nil, 0, false
}

func hasOnlyAcceptedProcedureTerminatorRecovery(root *tree_sitter.Node, problem ast.Range, tokenStart, tokenEnd int) bool {
	if root == nil {
		return false
	}
	foundProblem := false
	foundToken := false
	independent := false
	var walk func(*tree_sitter.Node)
	walk = func(node *tree_sitter.Node) {
		if node == nil {
			return
		}
		if node.IsError() || node.IsMissing() {
			range_ := ast.NodeRange(node)
			switch {
			case range_.StartByte == problem.StartByte && range_.EndByte == problem.EndByte:
				foundProblem = true
			case range_.StartByte == tokenStart && range_.EndByte == tokenEnd:
				foundToken = true
			default:
				independent = true
			}
		}
		for index := uint(0); index < node.ChildCount(); index++ {
			walk(node.Child(index))
		}
	}
	walk(root)
	return foundProblem && foundToken && !independent
}
