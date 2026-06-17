package ast

import (
	"os"

	tree_sitter_vba "github.com/harumiWeb/tree-sitter-vba/bindings/go"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

type ParseResult struct {
	Path       string
	Source     []byte
	Tree       *tree_sitter.Tree
	Root       *tree_sitter.Node
	HasError   bool
	HasMissing bool
}

type Range struct {
	StartLine   int `json:"startLine"`
	StartColumn int `json:"startColumn"`
	EndLine     int `json:"endLine"`
	EndColumn   int `json:"endColumn"`
	StartByte   int `json:"startByte"`
	EndByte     int `json:"endByte"`
}

type Parser struct {
	parser *tree_sitter.Parser
}

func NewParser() (*Parser, error) {
	parser := tree_sitter.NewParser()
	if err := parser.SetLanguage(tree_sitter.NewLanguage(tree_sitter_vba.Language())); err != nil {
		parser.Close()
		return nil, err
	}
	return &Parser{parser: parser}, nil
}

func (p *Parser) Close() {
	if p == nil || p.parser == nil {
		return
	}
	p.parser.Close()
	p.parser = nil
}

func (p *Parser) ParseFile(path string) (*ParseResult, error) {
	source, err := os.ReadFile(path)
	if err != nil {
		return nil, err
	}
	return p.Parse(path, source), nil
}

func (p *Parser) Parse(path string, source []byte) *ParseResult {
	tree := p.parser.Parse(source, nil)
	root := tree.RootNode()
	return &ParseResult{
		Path:       path,
		Source:     source,
		Tree:       tree,
		Root:       root,
		HasError:   root.HasError(),
		HasMissing: HasMissing(root),
	}
}

func (r *ParseResult) Close() {
	if r == nil || r.Tree == nil {
		return
	}
	r.Tree.Close()
	r.Tree = nil
	r.Root = nil
}

func NodeRange(node *tree_sitter.Node) Range {
	start := node.StartPosition()
	end := node.EndPosition()
	return Range{
		StartLine:   int(start.Row) + 1,
		StartColumn: int(start.Column) + 1,
		EndLine:     int(end.Row) + 1,
		EndColumn:   int(end.Column) + 1,
		StartByte:   int(node.StartByte()),
		EndByte:     int(node.EndByte()),
	}
}

func HasMissing(node *tree_sitter.Node) bool {
	if node == nil {
		return false
	}
	if node.IsMissing() {
		return true
	}
	for i := uint(0); i < node.ChildCount(); i++ {
		if HasMissing(node.Child(i)) {
			return true
		}
	}
	return false
}

func Walk(node *tree_sitter.Node, visit func(*tree_sitter.Node) bool) {
	if node == nil {
		return
	}
	if !visit(node) {
		return
	}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		Walk(node.NamedChild(i), visit)
	}
}
