package lint

import (
	"bytes"
	"context"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

// The diagnostic intentionally ignores writes. Deletion needs a stronger proof:
// exactly one syntactic occurrence of the name in its containing procedure.
// Count identifiers once per procedure, including writes and qualified names.
// Ambiguous occurrences conservatively withhold an action.
func (l Linter) unusedLocalQuickFixesContext(ctx context.Context, doc *vbaast.ParsedDocument) ([]QuickFix, error) {
	type candidate struct {
		name string
		line int
		edit FixEdit
	}
	var candidates []candidate
	var path, source string
	err := doc.ReadContext(ctx, func(view vbaast.ParsedView) error {
		if view.Root == nil || view.Root.HasError() || vbaast.HasMissing(view.Root) {
			return nil
		}
		path = view.Path
		var visit func(*tree_sitter.Node) error
		visit = func(node *tree_sitter.Node) error {
			if err := ctx.Err(); err != nil {
				return err
			}
			if strings.Contains(node.Kind(), "preprocessor") {
				return nil
			}
			if maintainabilityProcedureKind(node.Kind()) {
				header, _, _ := strings.Cut(node.Utf8Text(view.Source), "\n")
				for _, word := range strings.Fields(header) {
					if strings.EqualFold(word, "Static") {
						return nil
					}
				}
				counts := make(map[string]int)
				var local []candidate
				conditional := false
				var scan func(*tree_sitter.Node) error
				scan = func(child *tree_sitter.Node) error {
					if err := ctx.Err(); err != nil {
						return err
					}
					if strings.Contains(child.Kind(), "preprocessor") {
						conditional = true
					}
					if child.Kind() == "identifier" {
						counts[strings.ToLower(cleanIdentifier(child.Utf8Text(view.Source)))]++
					}
					if name, edit, ok := unusedLocalDimEdit(child, view.Source); ok {
						local = append(local, candidate{name: name, line: int(child.StartPosition().Row) + 1, edit: edit})
					}
					for i := range child.NamedChildCount() {
						if err := scan(child.NamedChild(i)); err != nil {
							return err
						}
					}
					return nil
				}
				if err := scan(node); err != nil {
					return err
				}
				if !conditional {
					for _, c := range local {
						if counts[strings.ToLower(c.name)] == 1 {
							candidates = append(candidates, c)
						}
					}
				}
				return nil
			}
			for i := range node.NamedChildCount() {
				if err := visit(node.NamedChild(i)); err != nil {
					return err
				}
			}
			return nil
		}
		if err := visit(view.Root); err != nil {
			return err
		}
		if len(candidates) > 0 {
			source = string(view.Source)
		}
		return nil
	})
	if err != nil || len(candidates) == 0 {
		return nil, err
	}
	// Borrow syntax separately, so the extractor never nests a document lease.
	file, err := symbols.InspectParsedContext(ctx, symbols.SourceOptions{RootDir: l.RootDir, IncludePrivate: true}, doc)
	if err != nil {
		return nil, err
	}
	issues, err := l.unusedLocalVariableIssuesContext(ctx, path, source, file)
	if err != nil {
		return nil, err
	}
	byLine := make(map[int]Issue, len(issues))
	for _, issue := range issues {
		byLine[issue.Line] = issue
	}
	var fixes []QuickFix
	for _, c := range candidates {
		if issue, ok := byLine[c.line]; ok && strings.EqualFold(issue.Symbol, c.name) {
			fixes = append(fixes, QuickFix{Title: "Remove unused local variable", Diagnostic: issue, Edits: []FixEdit{c.edit}})
		}
	}
	return fixes, ctx.Err()
}

func unusedLocalDimEdit(node *tree_sitter.Node, source []byte) (string, FixEdit, bool) {
	if node.Kind() != "variable_declaration" {
		return "", FixEdit{}, false
	}
	// The CST establishes declaration ownership; the spelling allowlist excludes
	// arrays, objects, implicit types, suffixes, fixed strings and continuations.
	text := node.Utf8Text(source)
	words := strings.Fields(text)
	if len(words) != 4 || !strings.EqualFold(words[0], "Dim") || !strings.EqualFold(words[2], "As") || strings.ContainsAny(text, "\r\n") {
		return "", FixEdit{}, false
	}
	switch strings.ToLower(words[3]) {
	case "boolean", "byte", "integer", "long", "longlong", "longptr", "single", "double", "currency", "date", "string", "variant":
	default:
		return "", FixEdit{}, false
	}
	var declarator *tree_sitter.Node
	for i := range node.NamedChildCount() {
		if child := node.NamedChild(i); child.Kind() == "variable_declarator" {
			if declarator != nil {
				return "", FixEdit{}, false
			}
			declarator = child
		}
	}
	if declarator == nil {
		return "", FixEdit{}, false
	}
	name := declarator.ChildByFieldName("name")
	if name == nil || name.Kind() != "identifier" || name.Utf8Text(source) != words[1] || cleanIdentifier(words[1]) != words[1] {
		return "", FixEdit{}, false
	}
	start, end := int(node.StartByte()), int(node.EndByte())
	lineStart := bytes.LastIndexAny(source[:start], "\r\n") + 1
	if len(bytes.Trim(source[lineStart:start], " \t")) != 0 {
		return "", FixEdit{}, false
	}
	lineEnd := end
	for lineEnd < len(source) && source[lineEnd] != '\r' && source[lineEnd] != '\n' {
		lineEnd++
	}
	suffix := strings.Trim(string(source[end:lineEnd]), " \t")
	if suffix != "" && !strings.HasPrefix(suffix, "'") && !startsWithKeyword(suffix, "Rem") {
		return "", FixEdit{}, false
	}
	return words[1], FixEdit{StartByte: start, EndByte: end}, true
}
