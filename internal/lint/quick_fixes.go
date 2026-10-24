package lint

import (
	"bytes"
	"context"
	"fmt"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

// FixEdit is a replacement in the original UTF-8 source, never in the result of
// a previous edit. Edits within one QuickFix must not overlap.
type FixEdit struct {
	StartByte int
	EndByte   int
	NewText   string
}

// QuickFix describes an explicit document-local correction and its diagnostic.
type QuickFix struct {
	Title      string
	Diagnostic Issue
	Edits      []FixEdit
}

// QuickFixesParsedContext scans only enabled correction providers. It does not
// run complete lint or construct ProcedureIR/CFG. Suppression is applied by the
// caller against the same source revision used to generate the edits.
func (l Linter) QuickFixesParsedContext(ctx context.Context, doc *vbaast.ParsedDocument) ([]QuickFix, error) {
	if !l.QuickFixesEnabled() {
		return nil, nil
	}
	var fixes []QuickFix
	err := doc.ReadContext(ctx, func(view vbaast.ParsedView) error {
		if l.Config.Lint.ProcedureNameConstant.Enabled {
			candidates, err := procedureNameConstantFixesContext(ctx, view.Root, view.Source, l.Config.Lint.ProcedureNameConstant)
			if err != nil {
				return err
			}
			for _, fix := range candidates {
				fixes = append(fixes, QuickFix{
					Title:      fmt.Sprintf("Update %s to %q", fix.ConstantName, fix.ExpectedName),
					Diagnostic: l.procedureNameConstantIssue(view.Path, view.Source, fix),
					Edits: []FixEdit{{StartByte: fix.StartByte, EndByte: fix.EndByte,
						NewText: `"` + strings.ReplaceAll(fix.ExpectedName, `"`, `""`) + `"`}},
				})
			}
		}
		if view.Root != nil && !view.Root.HasError() && !vbaast.HasMissing(view.Root) {
			hasOptionExplicit := false
			scanProcedures := l.Config.Lint.DetectLetAssignment || l.Config.Lint.DetectRedundantStepOne
			var walk func(*tree_sitter.Node) error
			walk = func(node *tree_sitter.Node) error {
				if err := ctx.Err(); err != nil {
					return err
				}
				// The grammar can represent an Option inside a preprocessor block
				// as a call. Treat that spelling as ambiguous, without changing lint.
				if l.Config.Lint.RequireOptionExplicit && !hasOptionExplicit && optionExplicitMayBePresent(node, view.Source) {
					hasOptionExplicit = true
				}
				if !scanProcedures && maintainabilityProcedureKind(node.Kind()) {
					return nil
				}
				if step := redundantStepOne(node, view.Source); l.Config.Lint.DetectRedundantStepOne && step != nil {
					if edit, ok := redundantStepOneEdit(node, step, view.Source); ok {
						fixes = append(fixes, QuickFix{
							Title:      "Remove redundant Step 1",
							Diagnostic: l.issueAt(view.Path, vbaast.NodeRange(step), "VB084", "information", "Step 1 restates the default step."),
							Edits:      []FixEdit{edit},
						})
					}
				}
				if l.Config.Lint.DetectLetAssignment && node.Kind() == "let_statement" {
					if edit, ok := explicitLetEdit(node, view.Source); ok {
						fixes = append(fixes, QuickFix{
							Title:      "Remove optional Let",
							Diagnostic: l.issueAt(view.Path, vbaast.NodeRange(node), "VB080", "information", "Explicit Let keyword is optional."),
							Edits:      []FixEdit{edit},
						})
					}
				}
				if l.Config.Lint.DetectRedundantOptionBaseZero && redundantOptionBaseZero(node, view.Source) {
					// Do not remove physical continuations or a multiline statement.
					if text := node.Utf8Text(view.Source); !strings.ContainsAny(text, "\r\n") {
						fixes = append(fixes, QuickFix{
							Title:      "Remove redundant Option Base 0",
							Diagnostic: l.issueAt(view.Path, vbaast.NodeRange(node), "VB086", "information", "Option Base 0 restates the default."),
							Edits:      []FixEdit{{StartByte: int(node.StartByte()), EndByte: int(node.EndByte())}},
						})
					}
				}
				for i := range node.NamedChildCount() {
					if err := walk(node.NamedChild(i)); err != nil {
						return err
					}
				}
				return nil
			}
			if err := walk(view.Root); err != nil {
				return err
			}
			// Any conditional Option Explicit makes absence ambiguous. Withhold even
			// when lint ignores a known-inactive branch, avoiding duplicate options.
			if l.Config.Lint.RequireOptionExplicit && !hasOptionExplicit {
				if edit, ok := missingOptionExplicitEdit(view.Root, view.Source); ok {
					fixes = append(fixes, QuickFix{
						Title:      "Add Option Explicit",
						Diagnostic: l.issue(view.Path, 1, "VB001", "warning", "Missing Option Explicit."),
						Edits:      []FixEdit{edit},
					})
				}
			}
		}
		return ctx.Err()
	})
	return fixes, err
}

func optionExplicitStatement(node *tree_sitter.Node, source []byte) bool {
	return node != nil && node.Kind() == "option_statement" && strings.EqualFold(normalizedNodeText(node, source), "Option Explicit")
}

func optionExplicitMayBePresent(node *tree_sitter.Node, source []byte) bool {
	if node.Kind() == "option_statement" {
		// Includes continued spellings, which this provider must not duplicate
		// even if the legacy diagnostic's normalized-text check misses them.
		return strings.Contains(strings.ToLower(node.Utf8Text(source)), "explicit")
	}
	if node.Kind() == "call_statement" {
		callee := node.ChildByFieldName("callee")
		return callee != nil && strings.EqualFold(callee.Utf8Text(source), "Option")
	}
	return false
}

func missingOptionExplicitEdit(root *tree_sitter.Node, source []byte) (FixEdit, bool) {
	boundary := 0
	pastHeader := false
	for i := range root.NamedChildCount() {
		node := root.NamedChild(i)
		switch node.Kind() {
		case "newline", "comment":
			continue
		case "frm_version_statement", "frm_begin_block", "frm_begin_property_block", "attribute_statement":
			if pastHeader {
				return FixEdit{}, false
			}
			start, end := int(node.StartByte()), int(node.EndByte())
			lineStart := bytes.LastIndexAny(source[:start], "\r\n") + 1
			if len(bytes.Trim(source[lineStart:start], " \t")) != 0 {
				return FixEdit{}, false
			}
			boundary = end
			if end > 0 && (source[end-1] == '\n' || source[end-1] == '\r') {
				continue
			}
			for boundary < len(source) && source[boundary] != '\r' && source[boundary] != '\n' {
				boundary++
			}
			suffix := strings.Trim(string(source[end:boundary]), " \t")
			if suffix != "" && !strings.HasPrefix(suffix, "'") && !startsWithKeyword(suffix, "Rem") {
				return FixEdit{}, false
			}
			if boundary < len(source) {
				if source[boundary] == '\r' && boundary+1 < len(source) && source[boundary+1] == '\n' {
					boundary++
				}
				boundary++
			}
		default:
			pastHeader = true
		}
	}
	// Insert before comments following the export preamble, so procedure or
	// module documentation remains attached to its original declaration.
	eol := "\n"
	if first := bytes.IndexAny(source, "\r\n"); first >= 0 && source[first] == '\r' {
		eol = "\r"
		if first+1 < len(source) && source[first+1] == '\n' {
			eol = "\r\n"
		}
	}
	text := "Option Explicit" + eol
	if boundary > 0 && source[boundary-1] != '\r' && source[boundary-1] != '\n' {
		text = eol + text
	}
	return FixEdit{StartByte: boundary, EndByte: boundary, NewText: text}, true
}

// QuickFixesEnabled lets adapters skip parsing when every provider is disabled.
func (l Linter) QuickFixesEnabled() bool {
	cfg := l.Config.Lint
	return cfg.RequireOptionExplicit || cfg.ProcedureNameConstant.Enabled || cfg.DetectRedundantStepOne || cfg.DetectLetAssignment || cfg.DetectRedundantOptionBaseZero
}

func explicitLetEdit(node *tree_sitter.Node, source []byte) (FixEdit, bool) {
	left := node.ChildByFieldName("left")
	if left == nil || left.StartByte() <= node.StartByte() {
		return FixEdit{}, false
	}
	prefix := string(source[node.StartByte():left.StartByte()])
	if len(prefix) <= len("Let") || !strings.EqualFold(strings.Trim(prefix, " \t"), "Let") {
		return FixEdit{}, false
	}
	return FixEdit{StartByte: int(node.StartByte()), EndByte: int(left.StartByte())}, true
}

func redundantStepOneEdit(loop, step *tree_sitter.Node, source []byte) (FixEdit, bool) {
	limit := loop.ChildByFieldName("end")
	if limit == nil || limit.EndByte() >= step.StartByte() {
		return FixEdit{}, false
	}
	// The grammar's external keyword scanner does not expose Step as a child.
	// Verify the entire CST-delimited gap, accepting horizontal trivia only.
	// In particular, never remove a continuation or an intervening comment.
	gap := string(source[limit.EndByte():step.StartByte()])
	if !strings.EqualFold(strings.Trim(gap, " \t"), "Step") {
		return FixEdit{}, false
	}
	return FixEdit{StartByte: int(limit.EndByte()), EndByte: int(step.EndByte())}, true
}
