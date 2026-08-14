package lint

import (
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
)

// conditionalBranchDiagnosticCode is reserved for compile-equivalent
// conditional-branch syntax.  The syntax pass stays independent from the
// linter's precedence stage so callers can decide how to suppress an
// overlapping parser-recovery finding (VB014).
const conditionalBranchDiagnosticCode = "VB062"

// conditionalBranchSyntaxIssues performs a conservative lexical pass over
// source statements and reports only branch forms whose invalidity is
// unambiguous from the source itself.  It intentionally uses blockStatements,
// the same continuation/colon splitter used by the VB014 structural recovery
// scanner.  Conditional-compilation directives, malformed continuations, or
// an otherwise ambiguous block boundary make the pass fail open and return no
// findings; the caller can then retain generic parser recovery.
//
// The method is intentionally independent of the AST.  The parser can recover
// an incomplete If/Else chain in several shapes, while the statement scanner
// still gives us stable source locations for the small set of high-confidence
// diagnostics this rule owns.
func (l Linter) conditionalBranchSyntaxIssues(path, source string) []Issue {
	statements, reliable := blockStatements(source)
	if !reliable || len(statements) == 0 {
		return nil
	}

	// A one-line If may contain an Else on the same physical/logical line.  It
	// does not open a multiline block, so branch ownership cannot be inferred
	// from the block stack for that group.  Keep those groups out of the orphan
	// branch check rather than guessing at VBA's one-line grammar.
	groupCounts := make(map[int]int)
	for _, statement := range statements {
		groupCounts[statement.group]++
	}
	inlineIfGroups := make(map[int]bool)
	invalidIfGroups := make(map[int]bool)
	for _, statement := range statements {
		text := conditionalBranchStatementText(statement.text)
		word, rest := firstVBAWord(text)
		if !strings.EqualFold(word, "if") || !hasVBAThenToken(rest) {
			continue
		}
		if groupCounts[statement.group] > 1 || conditionalBranchHasTextAfterThen(rest) {
			inlineIfGroups[statement.group] = true
		}
	}

	stack := make([]openBlock, 0, 4)
	issues := make([]Issue, 0)
	for index := 0; index < len(statements); index++ {
		statement := statements[index]
		text := conditionalBranchStatementText(statement.text)
		if text == "" {
			continue
		}
		if isRemComment(text) {
			// Rem comments consume the rest of a colon-separated physical line,
			// just like the structural recovery scanner.  Without this guard,
			// `Rem: Else` would be mistaken for an orphan branch.
			for index+1 < len(statements) && statements[index+1].group == statement.group {
				index++
			}
			continue
		}
		lower := strings.ToLower(strings.TrimSpace(text))

		word, rest := firstVBAWord(text)
		switch strings.ToLower(word) {
		case "if":
			if !hasVBAThenToken(rest) {
				issues = append(issues, l.conditionalBranchIssue(path, statement, "missing_then", "If", "An If statement must include Then.", "Add Then after the If condition."))
				if groupCounts[statement.group] > 1 {
					invalidIfGroups[statement.group] = true
				}
				// Keep a placeholder owner for following Else/ElseIf tokens.
				// The missing Then is the high-confidence defect; reporting
				// every branch that follows it as orphaned would be a cascade
				// from the same malformed If statement.
				if !inlineIfGroups[statement.group] {
					stack = append(stack, openBlock{kind: "if", closer: "End If", branchUncertain: true})
				}
			}
		case "elseif":
			if !hasVBAThenToken(rest) {
				if !invalidIfGroups[statement.group] {
					issues = append(issues, l.conditionalBranchIssue(path, statement, "missing_then", "ElseIf", "An ElseIf statement must include Then.", "Add Then after the ElseIf condition."))
				}
			} else if !inlineIfGroups[statement.group] && !invalidIfGroups[statement.group] {
				if len(stack) == 0 || stack[len(stack)-1].kind != "if" {
					issues = append(issues, l.conditionalBranchIssue(path, statement, "elseif_without_if", "ElseIf", "An ElseIf statement must belong to an If block.", "Move ElseIf inside a matching multiline If block."))
				} else if !stack[len(stack)-1].branchUncertain && stack[len(stack)-1].elseSeen {
					issues = append(issues, l.conditionalBranchIssue(path, statement, "elseif_after_else", "ElseIf", "ElseIf cannot follow Else in the same If block.", "Place ElseIf branches before the Else branch."))
				}
			}
		case "else":
			if !inlineIfGroups[statement.group] && !invalidIfGroups[statement.group] {
				if len(stack) == 0 || stack[len(stack)-1].kind != "if" {
					issues = append(issues, l.conditionalBranchIssue(path, statement, "else_without_if", "Else", "An Else statement must belong to an If block.", "Move Else inside a matching multiline If block."))
				} else if stack[len(stack)-1].branchUncertain {
					// Attach the branch to the malformed If without adding a
					// second diagnostic for the same source defect.
					stack[len(stack)-1].elseSeen = true
				} else if stack[len(stack)-1].elseSeen {
					issues = append(issues, l.conditionalBranchIssue(path, statement, "duplicate_else", "Else", "An If block can contain only one Else branch.", "Remove the duplicate Else branch."))
				} else {
					stack[len(stack)-1].elseSeen = true
				}
			}
		}

		// A branch with no unambiguous multiline owner is still useful evidence
		// above.  Do not attempt structural recovery after an already-invalid
		// statement: mismatched closers make ownership ambiguous, and callers
		// should retain VB014 instead of receiving speculative VB062 findings.
		if isProcedureEndStatement(lower) {
			if len(stack) != 0 {
				return conditionalBranchCertainIssues(issues)
			}
			continue
		}
		if _, ok := procedureStart(text, statement.line); ok {
			if len(stack) != 0 {
				return conditionalBranchCertainIssues(issues)
			}
			continue
		}

		if closer, count, ok := blockCloser(lower); ok {
			if count <= 0 || len(stack) < count {
				return conditionalBranchCertainIssues(issues)
			}
			for i := 0; i < count; i++ {
				if stack[len(stack)-1].kind != closer {
					return conditionalBranchCertainIssues(issues)
				}
				stack = stack[:len(stack)-1]
			}
			continue
		}

		hasFollowing := index+1 < len(statements) && statements[index+1].group == statement.group
		if block, ok := blockOpener(lower, hasFollowing); ok {
			stack = append(stack, block)
		}
	}

	// An unclosed block does not prove anything about branch ordering: parser
	// recovery may have dropped the closer or a nested block.  Let VB014 own
	// that case, while retaining the definite branch findings collected above
	// only when the entire stack is balanced.
	if len(stack) != 0 {
		return conditionalBranchCertainIssues(issues)
	}
	return issues
}

// conditionalBranchCertainIssues keeps findings whose ownership is explicit
// at the branch itself when a separate malformed closer makes later ordering
// ambiguous. Missing Then and orphan branches do not depend on a future
// closer; duplicate/after-Else ordering does and is left to VB014 in that
// recovery shape.
func conditionalBranchCertainIssues(issues []Issue) []Issue {
	if len(issues) == 0 {
		return nil
	}
	filtered := make([]Issue, 0, len(issues))
	for _, issue := range issues {
		switch issue.Kind {
		case "missing_then", "else_without_if", "elseif_without_if":
			filtered = append(filtered, issue)
		}
	}
	return filtered
}

// conditionalBranchIssue creates a stable statement-sized range and tags the
// finding with a machine-readable kind.  It is kept as a method so callers can
// use the same path-relative file handling as every other lint diagnostic.
func (l Linter) conditionalBranchIssue(path string, statement blockStatement, kind, symbol, message, suggestion string) Issue {
	column := statement.column
	if column < 1 {
		column = 1
	}
	text := strings.TrimSpace(statement.text)
	if n := leadingLineNumberLength(text); n > 0 {
		remainder := text[n:]
		text = strings.TrimLeft(remainder, " \t")
		column += n + len(remainder) - len(text)
	}
	if text == "" {
		text = statement.text
	}
	rng := vbaast.Range{
		StartLine:   statement.line,
		StartColumn: column,
		EndLine:     statement.line,
		EndColumn:   column + len(text),
	}
	issue := l.issueAt(path, rng, conditionalBranchDiagnosticCode, "error", message)
	issue.EndLine = rng.EndLine
	issue.EndColumn = rng.EndColumn
	issue.Kind = kind
	issue.Symbol = symbol
	issue.Suggestion = suggestion
	return issue
}

// conditionalBranchStatementText removes an optional VBA line number before
// classifying a statement.  Line numbers are source syntax, not part of the
// If/Else keyword and are already handled this way by the existing linter
// scanners.
func conditionalBranchStatementText(text string) string {
	text = strings.TrimSpace(text)
	if n := leadingLineNumberLength(text); n > 0 {
		text = strings.TrimSpace(text[n:])
	}
	return text
}

// hasVBAThenToken finds a Then keyword outside string/date literals.  A
// simple suffix check would misclassify valid conditions such as
// `If value = "Then" Then` and would turn this high-confidence rule into a
// source-text heuristic.
func hasVBAThenToken(text string) bool {
	for _, token := range conditionalBranchTokens(text) {
		if strings.EqualFold(token, "then") {
			return true
		}
	}
	return false
}

func conditionalBranchHasTextAfterThen(text string) bool {
	inString := false
	inDate := false
	for i := 0; i < len(text); {
		switch text[i] {
		case '"':
			if inString && i+1 < len(text) && text[i+1] == '"' {
				i += 2
				continue
			}
			inString = !inString
			i++
		case '#':
			if !inString {
				if inDate {
					inDate = false
				} else if startsDateLiteral(text, i) {
					inDate = true
				}
			}
			i++
		default:
			if inString || inDate || !isVBAIdentifierRune(rune(text[i])) {
				i++
				continue
			}
			start := i
			for i < len(text) && isVBAIdentifierRune(rune(text[i])) {
				i++
			}
			if strings.EqualFold(text[start:i], "then") {
				return strings.TrimSpace(text[i:]) != ""
			}
		}
	}
	return false
}

func conditionalBranchTokens(text string) []string {
	tokens := make([]string, 0, 4)
	inString := false
	inDate := false
	for i := 0; i < len(text); {
		switch text[i] {
		case '"':
			if inString && i+1 < len(text) && text[i+1] == '"' {
				i += 2
				continue
			}
			inString = !inString
			i++
		case '#':
			if !inString {
				if inDate {
					inDate = false
				} else if startsDateLiteral(text, i) {
					inDate = true
				}
			}
			i++
		default:
			if inString || inDate || !isVBAIdentifierRune(rune(text[i])) {
				i++
				continue
			}
			start := i
			for i < len(text) && isVBAIdentifierRune(rune(text[i])) {
				i++
			}
			tokens = append(tokens, text[start:i])
		}
	}
	return tokens
}
