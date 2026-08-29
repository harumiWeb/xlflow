package ast

import (
	"bytes"
	"context"
	"errors"
	"os"
	"runtime"
	"sync"
	"sync/atomic"

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
	mu       sync.Mutex
	treeGate chan struct{}
	result   *ParseResult
	readers  int
	closed   bool
}

// ParseDocument parses immutable source into a document-owned tree. The
// supplied source is copied so callers cannot mutate the bytes backing a
// shared parsed document after construction.
func ParseDocument(path string, source []byte) (*ParsedDocument, error) {
	return ParseDocumentContext(context.Background(), path, source)
}

// ParseDocumentContext parses immutable source while allowing tree-sitter to
// stop at a cooperative cancellation checkpoint.
func ParseDocumentContext(ctx context.Context, path string, source []byte) (*ParsedDocument, error) {
	return parseDocumentContext(ctx, path, source, nil)
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
	return parseDocumentContext(context.Background(), path, source, oldTree)
}

// ParseDocumentIncrementalIfAvailable performs an incremental parse only when
// the previous tree can be leased immediately. Interactive document updates
// must not wait behind a long-running reader of the obsolete revision.
func ParseDocumentIncrementalIfAvailable(path string, source, previousSource []byte, previous *ParsedDocument, edits []tree_sitter.InputEdit) (*ParsedDocument, error) {
	if previous == nil || len(edits) == 0 {
		return nil, ErrIncrementalParseUnavailable
	}
	oldTree, err := previous.tryCloneEditedTree(previousSource, edits)
	if err != nil {
		return nil, err
	}
	defer oldTree.Close()
	return parseDocumentContext(context.Background(), path, source, oldTree)
}

func parseDocumentContext(ctx context.Context, path string, source []byte, oldTree *tree_sitter.Tree) (*ParsedDocument, error) {
	if ctx == nil {
		ctx = context.Background()
	}
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	parser, err := NewParser()
	if err != nil {
		return nil, err
	}
	defer parser.Close()
	copySource := append([]byte(nil), source...)
	// go-tree-sitter v0.25.0 leaks the cgo handle for a non-nil
	// ParseOptions. Use the parser's cancellation flag with Parse (which passes
	// nil options) until a released binding is available, and join the watcher
	// before closing the parser so it cannot touch freed parser state.
	var parseDone chan struct{}
	var cancelDone chan struct{}
	if ctx.Done() != nil {
		cancellationFlag := new(uintptr)
		//nolint:staticcheck // ParseWithOptions leaks non-nil ParseOptions handles in v0.25.0.
		parser.parser.SetCancellationFlag(cancellationFlag)
		parseDone = make(chan struct{})
		cancelDone = make(chan struct{})
		go func() {
			defer close(cancelDone)
			select {
			case <-ctx.Done():
				atomic.StoreUintptr(cancellationFlag, 1)
			case <-parseDone:
			}
		}()
		defer func() {
			close(parseDone)
			<-cancelDone
			runtime.KeepAlive(cancellationFlag)
		}()
	}
	tree := parser.parser.Parse(copySource, oldTree)
	if tree == nil {
		if err := ctx.Err(); err != nil {
			return nil, err
		}
		return nil, ErrIncrementalParseUnavailable
	}
	if err := ctx.Err(); err != nil {
		tree.Close()
		return nil, err
	}
	root := tree.RootNode()
	if root == nil {
		tree.Close()
		return nil, ErrIncrementalParseUnavailable
	}
	treeGate := make(chan struct{}, 1)
	treeGate <- struct{}{}
	return &ParsedDocument{treeGate: treeGate, result: &ParseResult{
		Path: path, Source: copySource, Tree: tree, Root: root,
		HasError: root.HasError(), HasMissing: HasMissing(root),
	}}, nil
}

// Read serializes access to the document tree and invokes visit with its
// read-only view. Callers must finish all tree and node work before visit
// returns; retaining nodes beyond the callback is invalid.
func (d *ParsedDocument) Read(visit func(ParsedView) error) error {
	return d.ReadContext(context.Background(), visit)
}

// ReadContext is the cancellable form of Read. Cancellation while another
// reader owns the tree lease releases this caller's read reservation without
// waiting for the active tree walk to finish.
func (d *ParsedDocument) ReadContext(ctx context.Context, visit func(ParsedView) error) error {
	if d == nil {
		return ErrParsedDocumentClosed
	}
	if ctx == nil {
		ctx = context.Background()
	}
	d.mu.Lock()
	if d.closed || d.result == nil {
		d.mu.Unlock()
		return ErrParsedDocumentClosed
	}
	d.readers++
	result := d.result
	d.mu.Unlock()

	select {
	case <-ctx.Done():
		d.releaseRead()
		return ctx.Err()
	case <-d.treeGate:
	}
	defer func() { d.treeGate <- struct{}{} }()
	defer d.releaseRead()
	if err := ctx.Err(); err != nil {
		return err
	}
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

	<-d.treeGate
	defer func() { d.treeGate <- struct{}{} }()
	defer d.releaseRead()
	clone := result.Tree.Clone()
	for index := range edits {
		clone.Edit(&edits[index])
	}
	if clone == nil {
		return nil, ErrIncrementalParseUnavailable
	}
	return clone, nil
}

func (d *ParsedDocument) tryCloneEditedTree(expectedSource []byte, edits []tree_sitter.InputEdit) (*tree_sitter.Tree, error) {
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

	select {
	case <-d.treeGate:
	default:
		d.releaseRead()
		return nil, ErrIncrementalParseUnavailable
	}
	defer func() { d.treeGate <- struct{}{} }()
	defer d.releaseRead()
	if !bytes.Equal(result.Source, expectedSource) {
		return nil, ErrIncrementalParseUnavailable
	}
	clone := result.Tree.Clone()
	if clone == nil {
		return nil, ErrIncrementalParseUnavailable
	}
	for index := range edits {
		clone.Edit(&edits[index])
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
