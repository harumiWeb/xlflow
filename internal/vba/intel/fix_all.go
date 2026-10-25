package intel

import (
	"context"
	"fmt"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/lint"
)

const documentFixAllPrefix = "source.fixAll.xlflow."

// Eligibility is explicit and independent of fix_available. Intent-changing
// VB001 and the type-backed Set corrections are not document Fix All rules.
var documentFixAllRules = [...]string{"VB044", "VB080", "VB084", "VB086"}

// FixAllCodeActionKinds returns fresh rule-specific kinds in stable rule order.
func FixAllCodeActionKinds() []string {
	kinds := make([]string, len(documentFixAllRules))
	for i, code := range documentFixAllRules {
		kinds[i] = documentFixAllKind(code)
	}
	return kinds
}

func documentFixAllKind(code string) string {
	return documentFixAllPrefix + strings.ToLower(code)
}

func requestedFixAll(only []string) bool {
	for _, kind := range only {
		// An empty kind selects ordinary actions, not an expensive source scan.
		if kind == "" {
			continue
		}
		for _, supported := range FixAllCodeActionKinds() {
			if codeActionKindAllowed(supported, []string{kind}) {
				return true
			}
		}
	}
	return false
}

func (a Analyzer) documentFixAllCodeActionsContext(ctx context.Context, doc Document, only []string) ([]CodeAction, error) {
	// Root-kind requests must not broaden an explicitly selected source kind.
	only = slices.DeleteFunc(slices.Clone(only), func(kind string) bool { return kind == "" })
	// Disable unrequested/ineligible providers before borrowing syntax. In
	// particular, never run type-backed diagnostics or build a symbol index.
	cfg := a.Config
	cfg.Lint.RequireOptionExplicit = false
	cfg.Lint.ProcedureNameConstant.Enabled = cfg.Lint.ProcedureNameConstant.Enabled && codeActionKindAllowed(documentFixAllKind("VB044"), only)
	cfg.Lint.DetectLetAssignment = cfg.Lint.DetectLetAssignment && codeActionKindAllowed(documentFixAllKind("VB080"), only)
	cfg.Lint.DetectRedundantStepOne = cfg.Lint.DetectRedundantStepOne && codeActionKindAllowed(documentFixAllKind("VB084"), only)
	cfg.Lint.DetectRedundantOptionBaseZero = cfg.Lint.DetectRedundantOptionBaseZero && codeActionKindAllowed(documentFixAllKind("VB086"), only)
	linter := lint.Linter{RootDir: a.RootDir, Config: cfg}
	if !linter.QuickFixesEnabled() {
		return nil, ctx.Err()
	}
	parsed, closeParsed, err := parsedDocumentForDocument(doc)
	if err != nil {
		return nil, err
	}
	defer closeParsed()
	selection := Range{End: positionForDocumentByteOffset(doc, len(doc.Source))}
	fixes, err := a.lintQuickFixCodeActionsContext(ctx, doc, parsed, selection, linter)
	if err != nil {
		return nil, err
	}
	var out []CodeAction
	for _, code := range documentFixAllRules {
		if err := ctx.Err(); err != nil {
			return nil, err
		}
		if !codeActionKindAllowed(documentFixAllKind(code), only) {
			continue
		}
		if action, ok := aggregateDocumentFixes(doc, code, fixes); ok {
			out = append(out, action)
		}
	}
	return out, ctx.Err()
}

// aggregateDocumentFixes validates all ranges against the captured document.
// A conflict rejects the whole rule action, never a subset of its corrections.
func aggregateDocumentFixes(doc Document, code string, fixes []CodeAction) (CodeAction, bool) {
	if !slices.Contains(documentFixAllRules[:], code) {
		return CodeAction{}, false
	}
	action := CodeAction{Title: fmt.Sprintf("Fix all %s occurrences in this document", code), Kind: documentFixAllKind(code)}
	for _, fix := range fixes {
		if len(fix.Diagnostics) != 1 || fix.Diagnostics[0].Code != code {
			continue
		}
		if len(fix.Edits) == 0 {
			return CodeAction{}, false
		}
		for _, edit := range fix.Edits {
			if comparePosition(edit.Range.Start, edit.Range.End) > 0 || !validFixAllPosition(doc, edit.Range.Start) || !validFixAllPosition(doc, edit.Range.End) {
				return CodeAction{}, false
			}
			action.Edits = append(action.Edits, edit)
		}
		action.Diagnostics = append(action.Diagnostics, fix.Diagnostics[0])
	}
	if len(action.Edits) == 0 {
		return CodeAction{}, false
	}
	slices.SortFunc(action.Edits, func(left, right CodeActionEdit) int {
		if result := comparePosition(left.Range.Start, right.Range.Start); result != 0 {
			return result
		}
		return comparePosition(left.Range.End, right.Range.End)
	})
	for i := 1; i < len(action.Edits); i++ {
		previous, current := action.Edits[i-1].Range, action.Edits[i].Range
		if comparePosition(current.Start, previous.End) < 0 || current.Start == previous.Start {
			return CodeAction{}, false
		}
	}
	return action, true
}

func validFixAllPosition(doc Document, position Position) bool {
	if position.Line < 0 || position.Character < 0 {
		return false
	}
	offset := byteOffsetForDocumentPosition(doc, position)
	return positionForDocumentByteOffset(doc, offset) == position
}
