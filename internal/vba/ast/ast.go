package ast

import (
	"errors"
	"os"
	"sync"

	tree_sitter_vba "github.com/harumiWeb/tree-sitter-vba/bindings/go"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

// ErrParsedDocumentClosed reports an attempt to read a document whose owner
// has retired it. A parsed document never reopens after Close.
var ErrParsedDocumentClosed = errors.New("parsed VBA document is closed")

type ParseResult struct {
	Path       string
	Source     []byte
	Tree       *tree_sitter.Tree
	Root       *tree_sitter.Node
	HasError   bool
	HasMissing bool
}

// ParsedView is the read-only tree-sitter state exposed during a
// ParsedDocument.Read callback. Root and Source are valid only for the
// callback's duration and must not be retained or mutated by callers.
type ParsedView struct {
	Path       string
	Source     []byte
	Root       *tree_sitter.Node
	HasError   bool
	HasMissing bool
}

// ParsedDocument owns one ParseResult and its tree-sitter tree. ParseDocument
// creates and closes the parser; this type alone owns and closes the resulting
// tree. Tree-sitter trees are not thread safe, so Read serializes tree access
// for this document. Close is idempotent, rejects new reads, and releases the
// tree exactly once after every in-flight read callback returns.
type ParsedDocument struct {
	mu      sync.Mutex
	treeMu  sync.Mutex
	result  *ParseResult
	readers int
	closed  bool
}

// ParseDocument parses immutable source into a document-owned tree. The
// supplied source is copied so callers cannot mutate the bytes backing a
// shared parsed document after construction.
func ParseDocument(path string, source []byte) (*ParsedDocument, error) {
	parser, err := NewParser()
	if err != nil {
		return nil, err
	}
	defer parser.Close()
	return &ParsedDocument{result: parser.Parse(path, append([]byte(nil), source...))}, nil
}

// Read serializes access to the document tree and invokes visit with its
// read-only view. Callers must finish all tree and node work before visit
// returns; retaining nodes beyond the callback is invalid.
func (d *ParsedDocument) Read(visit func(ParsedView) error) error {
	if d == nil {
		return ErrParsedDocumentClosed
	}
	d.mu.Lock()
	if d.closed || d.result == nil {
		d.mu.Unlock()
		return ErrParsedDocumentClosed
	}
	d.readers++
	result := d.result
	d.mu.Unlock()

	d.treeMu.Lock()
	defer d.treeMu.Unlock()
	defer d.releaseRead()
	return visit(ParsedView{
		Path:       result.Path,
		Source:     result.Source,
		Root:       result.Root,
		HasError:   result.HasError,
		HasMissing: result.HasMissing,
	})
}

func (d *ParsedDocument) releaseRead() {
	d.mu.Lock()
	defer d.mu.Unlock()
	d.readers--
	if d.closed && d.readers == 0 {
		d.closeResultLocked()
	}
}

// Close retires the document. It is safe to call more than once and never
// closes a tree while a Read callback can still access it.
func (d *ParsedDocument) Close() {
	if d == nil {
		return
	}
	d.mu.Lock()
	defer d.mu.Unlock()
	d.closed = true
	if d.readers == 0 {
		d.closeResultLocked()
	}
}

func (d *ParsedDocument) closeResultLocked() {
	if d.result == nil {
		return
	}
	d.result.Close()
	d.result = nil
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
