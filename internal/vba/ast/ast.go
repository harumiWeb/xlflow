package ast

import (
	"bytes"
	"errors"
	"os"
	"sync"

	tree_sitter_vba "github.com/harumiWeb/tree-sitter-vba/bindings/go"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

// ErrParsedDocumentClosed reports an attempt to read a document whose owner
// has retired it. A parsed document never reopens after Close.
var ErrParsedDocumentClosed = errors.New("parsed VBA document is closed")

// ErrIncrementalParseUnavailable reports that a previous parsed document can
// no longer safely provide a tree for incremental parsing.
var ErrIncrementalParseUnavailable = errors.New("incremental VBA parse is unavailable")

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
	return parseDocument(path, source, nil)
}

// ParseDocumentIncremental parses source using an edited clone of previous's
// tree. It never mutates the previous document's tree, so a published
// immutable analysis snapshot remains readable while its successor is built.
// The caller must fall back to ParseDocument when this returns
// ErrIncrementalParseUnavailable.
func ParseDocumentIncremental(path string, source []byte, previous *ParsedDocument, edits []tree_sitter.InputEdit) (*ParsedDocument, error) {
	if previous == nil || len(edits) == 0 {
		return nil, ErrIncrementalParseUnavailable
	}
	oldTree, err := previous.cloneEditedTree(edits)
	if err != nil {
		return nil, err
	}
	defer oldTree.Close()
	return parseDocument(path, source, oldTree)
}

func parseDocument(path string, source []byte, oldTree *tree_sitter.Tree) (*ParsedDocument, error) {
	parser, err := NewParser()
	if err != nil {
		return nil, err
	}
	defer parser.Close()
	copySource := append([]byte(nil), source...)
	tree := parser.parser.Parse(copySource, oldTree)
	if tree == nil {
		return nil, ErrIncrementalParseUnavailable
	}
	root := tree.RootNode()
	if root == nil {
		tree.Close()
		return nil, ErrIncrementalParseUnavailable
	}
	return &ParsedDocument{result: &ParseResult{
		Path: path, Source: copySource, Tree: tree, Root: root,
		HasError: root.HasError(), HasMissing: HasMissing(root),
	}}, nil
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

// SourceMatches reports whether this document still owns exactly source. It
// takes the same read lease as Read so a concurrent Close cannot release the
// tree during the comparison.
func (d *ParsedDocument) SourceMatches(source []byte) bool {
	matched := false
	if d == nil {
		return false
	}
	if d.Read(func(view ParsedView) error {
		matched = bytes.Equal(view.Source, source)
		return nil
	}) != nil {
		return false
	}
	return matched
}

func (d *ParsedDocument) releaseRead() {
	d.mu.Lock()
	defer d.mu.Unlock()
	d.readers--
	if d.closed && d.readers == 0 {
		d.closeResultLocked()
	}
}

// cloneEditedTree takes a read lease while it clones and edits the tree. This
// makes Close wait until the clone is complete and serializes the operation
// with every tree reader.
func (d *ParsedDocument) cloneEditedTree(edits []tree_sitter.InputEdit) (*tree_sitter.Tree, error) {
	if d == nil {
		return nil, ErrIncrementalParseUnavailable
	}
	d.mu.Lock()
	if d.closed || d.result == nil || d.result.Tree == nil {
		d.mu.Unlock()
		return nil, ErrIncrementalParseUnavailable
	}
	d.readers++
	result := d.result
	d.mu.Unlock()

	d.treeMu.Lock()
	clone := result.Tree.Clone()
	for index := range edits {
		clone.Edit(&edits[index])
	}
	d.treeMu.Unlock()
	d.releaseRead()
	if clone == nil {
		return nil, ErrIncrementalParseUnavailable
	}
	return clone, nil
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
