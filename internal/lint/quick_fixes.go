package lint

import (
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
	if !l.Config.Lint.ProcedureNameConstant.Enabled && !l.Config.Lint.DetectRedundantStepOne {
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
		if l.Config.Lint.DetectRedundantStepOne && view.Root != nil && !view.Root.HasError() && !vbaast.HasMissing(view.Root) {
			var walk func(*tree_sitter.Node) error
			walk = func(node *tree_sitter.Node) error {
				if err := ctx.Err(); err != nil {
					return err
				}
				if step := redundantStepOne(node, view.Source); step != nil {
					if edit, ok := redundantStepOneEdit(node, step, view.Source); ok {
						fixes = append(fixes, QuickFix{
							Title:      "Remove redundant Step 1",
							Diagnostic: l.issueAt(view.Path, vbaast.NodeRange(step), "VB084", "information", "Step 1 restates the default step."),
							Edits:      []FixEdit{edit},
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
			return walk(view.Root)
		}
		return ctx.Err()
	})
	return fixes, err
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
