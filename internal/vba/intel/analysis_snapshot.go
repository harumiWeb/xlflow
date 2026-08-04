package intel

import (
	"context"
	"crypto/sha256"
	"encoding/hex"
	"errors"
	"sort"
	"strings"
	"sync"
	"sync/atomic"
	"unsafe"

	"github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/calls"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

var errAnalysisSnapshotRetired = errors.New("analysis snapshot is retired")

// ErrIncrementalSnapshotUnavailable reports that a successor cannot safely
// reuse the preceding snapshot's tree and must be parsed from scratch.
var ErrIncrementalSnapshotUnavailable = errors.New("incremental analysis snapshot is unavailable")

// ProcedureInfo describes the source range occupied by a VBA procedure.
type ProcedureInfo struct {
	Name  string
	Range Range
}

// AnalysisSnapshot is the immutable source state for one document revision.
// Derived Go values are initialized once and may be read concurrently. Its
// lazily-created ParsedDocument owns the revision's tree-sitter result; tree
// access is serialized by ParsedDocument because tree-sitter trees are not
// thread safe. Retire closes that document after its in-flight readers finish.
type AnalysisSnapshot struct {
	uri        string
	path       string
	version    int32
	moduleKind string
	source     string
	sourceHash [sha256.Size]byte
	lines      []string
	lineStarts []int
	lineEnds   []int

	proceduresOnce sync.Once
	procedures     []ProcedureInfo
	procedureLines []int

	symbolsMu   sync.Mutex
	symbolsWait chan struct{}
	symbolsDone bool
	symbols     []Symbol
	symbolsErr  error

	callSitesOnce sync.Once
	callSites     calls.FileResult
	callSitesErr  error

	procedureIRMu   sync.Mutex
	procedureIRWait chan struct{}
	procedureIRDone bool
	procedureIR     procedureir.DocumentIR
	procedureIRErr  error

	controlFlowMu   sync.Mutex
	controlFlowWait chan struct{}
	controlFlowDone bool
	controlFlow     vbacfg.Document
	controlFlowErr  error

	indexOnce sync.Once
	index     *documentIndex
	indexErr  error

	semanticOnce        sync.Once
	semanticIdentifiers [][]byteSpan

	parsedMu              sync.Mutex
	parsedDocument        *ast.ParsedDocument
	parsedErr             error
	parseDocument         func(string, []byte) (*ast.ParsedDocument, error)
	parseCount            atomic.Uint64
	fullHashCount         atomic.Uint64
	procedureIRBuildCount atomic.Uint64
	procedureIRReuseCount atomic.Uint64
	cfgBuildCount         atomic.Uint64
	cfgReuseCount         atomic.Uint64
	artifacts             *procedureArtifactStore

	retired atomic.Bool
}

// NewAnalysisSnapshot captures doc as an immutable analysis revision.
func NewAnalysisSnapshot(doc Document) *AnalysisSnapshot {
	source := doc.Source
	normalized := strings.ReplaceAll(source, "\r\n", "\n")
	normalized = strings.ReplaceAll(normalized, "\r", "\n")
	lineStarts, lineEnds := buildSourceLineMap(source)
	return &AnalysisSnapshot{
		uri:           doc.URI,
		path:          doc.Path,
		version:       doc.Version,
		moduleKind:    doc.ModuleKind,
		source:        source,
		sourceHash:    sha256.Sum256([]byte(source)),
		lines:         strings.Split(normalized, "\n"),
		lineStarts:    lineStarts,
		lineEnds:      lineEnds,
		parseDocument: ast.ParseDocument,
		artifacts:     newProcedureArtifactStore(),
	}
}

func buildSourceLineMap(source string) ([]int, []int) {
	starts := []int{0}
	ends := make([]int, 0, strings.Count(source, "\n")+1)
	for i := 0; i < len(source); i++ {
		switch source[i] {
		case '\n':
			ends = append(ends, i)
			starts = append(starts, i+1)
		case '\r':
			ends = append(ends, i)
			if i+1 < len(source) && source[i+1] == '\n' {
				i++
			}
			starts = append(starts, i+1)
		}
	}
	ends = append(ends, len(source))
	return starts, ends
}

func (s *AnalysisSnapshot) byteOffset(pos Position) int {
	if s == nil || pos.Line < 0 {
		return 0
	}
	if pos.Line >= len(s.lineStarts) {
		return len(s.source)
	}
	start, end := s.lineStarts[pos.Line], s.lineEnds[pos.Line]
	index := byteIndexForUTF16(s.source[start:end], pos.Character)
	return start + min(index, end-start)
}

func (s *AnalysisSnapshot) position(offset int) Position {
	if s == nil || offset <= 0 {
		return Position{}
	}
	if offset >= len(s.source) {
		line := len(s.lineStarts) - 1
		return Position{Line: line, Character: utf16Len(s.source[s.lineStarts[line]:s.lineEnds[line]])}
	}
	line := sort.Search(len(s.lineStarts), func(i int) bool { return s.lineStarts[i] > offset }) - 1
	line = max(0, line)
	end := min(offset, s.lineEnds[line])
	return Position{Line: line, Character: utf16Len(s.source[s.lineStarts[line]:end])}
}

func positionForSourceMap(source string, lineStarts, lineEnds []int, offset int) Position {
	if offset <= 0 || len(lineStarts) == 0 {
		return Position{}
	}
	if offset >= len(source) {
		line := len(lineStarts) - 1
		return Position{Line: line, Character: utf16Len(source[lineStarts[line]:lineEnds[line]])}
	}
	line := sort.Search(len(lineStarts), func(i int) bool { return lineStarts[i] > offset }) - 1
	line = max(0, line)
	end := min(offset, lineEnds[line])
	return Position{Line: line, Character: utf16Len(source[lineStarts[line]:end])}
}

// NewAnalysisSnapshotWithParsedDocument captures doc and seeds it with a
// parsed document produced while preparing this exact revision. Ownership of
// parsed transfers to the returned snapshot, which retires it exactly once.
func NewAnalysisSnapshotWithParsedDocument(doc Document, parsed *ast.ParsedDocument) *AnalysisSnapshot {
	snapshot := NewAnalysisSnapshot(doc)
	if parsed != nil {
		snapshot.parsedDocument = parsed
		snapshot.parseCount.Add(1)
	}
	return snapshot
}

// NewIncrementalAnalysisSnapshot creates doc from an edited clone of
// previous's tree. previous remains immutable and continues to own its tree;
// the returned snapshot owns only the newly parsed tree.
func NewIncrementalAnalysisSnapshot(doc Document, previous *AnalysisSnapshot, edits []tree_sitter.InputEdit) (*AnalysisSnapshot, error) {
	if previous == nil || len(edits) == 0 || previous.Retired() {
		return nil, ErrIncrementalSnapshotUnavailable
	}
	parsed, err := previous.ParsedDocument()
	if err != nil || parsed == nil || !parsed.SourceMatches([]byte(previous.Source())) {
		return nil, ErrIncrementalSnapshotUnavailable
	}
	next, err := ast.ParseDocumentIncremental(doc.Path, []byte(doc.Source), parsed, edits)
	if err != nil || next == nil {
		return nil, ErrIncrementalSnapshotUnavailable
	}
	snapshot := NewAnalysisSnapshotWithParsedDocument(doc, next)
	snapshot.artifacts = previous.artifacts.clone()
	return snapshot, nil
}

// NewSuccessorAnalysisSnapshotWithParsedDocument creates a full-replacement
// revision that may inherit only completed immutable procedure artifacts from
// the same open-document lifecycle.
func NewSuccessorAnalysisSnapshotWithParsedDocument(doc Document, parsed *ast.ParsedDocument, previous *AnalysisSnapshot) *AnalysisSnapshot {
	snapshot := NewAnalysisSnapshotWithParsedDocument(doc, parsed)
	if previous != nil && !previous.Retired() {
		snapshot.artifacts = previous.artifacts.clone()
	}
	return snapshot
}

// Document returns a document view associated with this snapshot.
func (s *AnalysisSnapshot) Document() Document {
	if s == nil {
		return Document{}
	}
	return Document{
		URI: s.uri, Path: s.path, Source: s.source, ModuleKind: s.moduleKind,
		Version: s.version, Snapshot: s,
	}
}

func (s *AnalysisSnapshot) URI() string        { return s.uri }
func (s *AnalysisSnapshot) Path() string       { return s.path }
func (s *AnalysisSnapshot) Version() int32     { return s.version }
func (s *AnalysisSnapshot) ModuleKind() string { return s.moduleKind }
func (s *AnalysisSnapshot) Source() string     { return s.source }

// SourceHash returns the lowercase hexadecimal SHA-256 source digest.
func (s *AnalysisSnapshot) SourceHash() string {
	if s == nil {
		return ""
	}
	return hex.EncodeToString(s.sourceHash[:])
}

func (s *AnalysisSnapshot) sameRevision(doc Document) bool {
	if s == nil || s.retired.Load() || s.uri != doc.URI || s.path != doc.Path ||
		s.version != doc.Version || s.moduleKind != doc.ModuleKind ||
		len(s.source) != len(doc.Source) {
		return false
	}
	// A Document returned by this snapshot retains the immutable source
	// string's backing storage. Comparing that identity avoids hashing the
	// entire source on every cache lookup while length keeps substring aliases
	// from being mistaken for this revision. Independently allocated strings
	// still use the digest so Matches preserves its value semantics.
	if unsafe.StringData(s.source) == unsafe.StringData(doc.Source) {
		return true
	}
	s.fullHashCount.Add(1)
	return s.sourceHash == sha256.Sum256([]byte(doc.Source))
}

// Matches reports whether doc describes the exact immutable revision captured by the snapshot.
func (s *AnalysisSnapshot) Matches(doc Document) bool { return s.sameRevision(doc) }

// Lines returns a defensive copy of normalized source lines.
func (s *AnalysisSnapshot) Lines() []string {
	if s == nil {
		return nil
	}
	return append([]string(nil), s.lines...)
}

func (s *AnalysisSnapshot) sourceLines() []string { return s.lines }

// Procedures returns a defensive copy of the procedure index.
func (s *AnalysisSnapshot) Procedures() []ProcedureInfo {
	if s == nil {
		return nil
	}
	s.initProcedures()
	return append([]ProcedureInfo(nil), s.procedures...)
}

func (s *AnalysisSnapshot) initProcedures() {
	s.proceduresOnce.Do(func() {
		s.procedureLines = make([]int, len(s.lines))
		for i := range s.procedureLines {
			s.procedureLines[i] = -1
		}
		depth, active := 0, -1
		for lineNo, line := range s.lines {
			text := strings.TrimSpace(line[:codeLimit(line)])
			if text == "" {
				continue
			}
			lower := strings.ToLower(text)
			switch {
			case strings.HasPrefix(lower, "end sub") || strings.HasPrefix(lower, "end function") || strings.HasPrefix(lower, "end property"):
				if depth > 0 {
					depth--
				}
				if depth == 0 && active >= 0 {
					s.procedures[active].Range.End = Position{Line: lineNo, Character: utf16Len(line)}
					active = -1
				}
			case procedureStartLine(lower):
				depth++
				if depth == 1 {
					if name := procedureNameFromLine(text); name != "" {
						s.procedures = append(s.procedures, ProcedureInfo{
							Name:  name,
							Range: Range{Start: Position{Line: lineNo}, End: Position{Line: len(s.lines)}},
						})
						active = len(s.procedures) - 1
					}
				}
			}
		}
		for index, procedure := range s.procedures {
			lastLine := min(procedure.Range.End.Line, len(s.procedureLines)-1)
			for lineNo := procedure.Range.Start.Line; lineNo <= lastLine && lineNo >= 0; lineNo++ {
				s.procedureLines[lineNo] = index
			}
		}
	})
}

func (s *AnalysisSnapshot) procedureAt(pos Position) (string, *Range) {
	if s == nil {
		return "", nil
	}
	s.initProcedures()
	if pos.Line < 0 || pos.Line >= len(s.procedureLines) {
		return "", nil
	}
	index := s.procedureLines[pos.Line]
	if index < 0 || index >= len(s.procedures) {
		return "", nil
	}
	procedure := s.procedures[index]
	if comparePosition(procedure.Range.Start, pos) > 0 || comparePosition(pos, procedure.Range.End) > 0 {
		return "", nil
	}
	rng := procedure.Range
	return procedure.Name, &rng
}

// SourceSymbols returns snapshot-scoped source symbols and whether the lazy value was already initialized.
func (s *AnalysisSnapshot) SourceSymbols(load DocumentSymbolLoader) ([]Symbol, bool, error) {
	return s.SourceSymbolsContext(context.Background(), func(context.Context) ([]Symbol, error) { return load() })
}

// SourceSymbolsContext is a retryable single-flight cache. Cancellation is
// scoped to the active build and is not retained as a revision-wide result.
func (s *AnalysisSnapshot) SourceSymbolsContext(ctx context.Context, load func(context.Context) ([]Symbol, error)) ([]Symbol, bool, error) {
	if s == nil {
		syms, err := load(ctx)
		return cloneAnalysisSymbols(syms), false, err
	}
	for {
		if err := ctx.Err(); err != nil {
			return nil, false, err
		}
		s.symbolsMu.Lock()
		if s.symbolsDone {
			result, err := cloneAnalysisSymbols(s.symbols), s.symbolsErr
			s.symbolsMu.Unlock()
			return result, true, err
		}
		if s.symbolsWait != nil {
			wait := s.symbolsWait
			s.symbolsMu.Unlock()
			select {
			case <-ctx.Done():
				return nil, false, ctx.Err()
			case <-wait:
			}
			continue
		}
		wait := make(chan struct{})
		s.symbolsWait = wait
		s.symbolsMu.Unlock()
		defer func() {
			s.symbolsMu.Lock()
			if s.symbolsWait == wait {
				s.symbolsWait = nil
				close(wait)
			}
			s.symbolsMu.Unlock()
		}()
		result, err := load(ctx)
		if err == nil {
			err = ctx.Err()
		}
		s.symbolsMu.Lock()
		if !isRetryableContextError(err) {
			s.symbols, s.symbolsErr, s.symbolsDone = cloneAnalysisSymbols(result), err, true
		}
		s.symbolsMu.Unlock()
		return cloneAnalysisSymbols(result), false, err
	}
}

// RawCallSites returns snapshot-scoped syntax-local call sites and whether the
// lazy value was already initialized. Both successful results and extraction
// errors are cached for this immutable revision. The returned result is a deep
// copy and can be modified by the caller.
func (s *AnalysisSnapshot) RawCallSites(load func() (calls.FileResult, error)) (calls.FileResult, bool, error) {
	if s == nil {
		result, err := load()
		return calls.CloneFileResult(result), false, err
	}
	initialized := true
	s.callSitesOnce.Do(func() {
		initialized = false
		result, err := load()
		s.callSites = calls.CloneFileResult(result)
		s.callSitesErr = err
	})
	return calls.CloneFileResult(s.callSites), initialized, s.callSitesErr
}

// ProcedureIR returns the snapshot-scoped procedure analysis IR and whether
// the lazy value was already initialized. Both successful results and build
// errors are cached for this immutable revision. The returned IR is a deep
// copy and can be modified by the caller.
func (s *AnalysisSnapshot) ProcedureIR(load func() (procedureir.DocumentIR, error)) (procedureir.DocumentIR, bool, error) {
	return s.ProcedureIRContext(context.Background(), func(context.Context) (procedureir.DocumentIR, error) { return load() })
}

// ProcedureIRContext is a retryable single-flight cache. Cancellation is a
// property of one build attempt, not of the immutable revision, so canceled
// attempts wake waiters without storing a value or permanent error.
func (s *AnalysisSnapshot) ProcedureIRContext(ctx context.Context, load func(context.Context) (procedureir.DocumentIR, error)) (procedureir.DocumentIR, bool, error) {
	if s == nil {
		result, err := load(ctx)
		return procedureir.Clone(result), false, err
	}
	for {
		if err := ctx.Err(); err != nil {
			return procedureir.DocumentIR{}, false, err
		}
		s.procedureIRMu.Lock()
		if s.procedureIRDone {
			result, err := procedureir.Clone(s.procedureIR), s.procedureIRErr
			s.procedureIRMu.Unlock()
			return result, true, err
		}
		if s.procedureIRWait != nil {
			wait := s.procedureIRWait
			s.procedureIRMu.Unlock()
			select {
			case <-ctx.Done():
				return procedureir.DocumentIR{}, false, ctx.Err()
			case <-wait:
			}
			continue
		}
		wait := make(chan struct{})
		s.procedureIRWait = wait
		s.procedureIRMu.Unlock()
		defer func() {
			s.procedureIRMu.Lock()
			if s.procedureIRWait == wait {
				s.procedureIRWait = nil
				close(wait)
			}
			s.procedureIRMu.Unlock()
		}()
		result, err := load(ctx)
		if err == nil {
			err = ctx.Err()
		}
		s.procedureIRMu.Lock()
		if !isRetryableContextError(err) {
			s.procedureIR, s.procedureIRErr, s.procedureIRDone = procedureir.Clone(result), err, true
		}
		s.procedureIRMu.Unlock()
		if err == nil {
			s.seedProcedureArtifacts(result)
		}
		return procedureir.Clone(result), false, err
	}
}

// ControlFlowGraphs returns the snapshot-scoped procedure control-flow graphs
// and whether the lazy value was already initialized. Both successful results
// and build errors are cached for this immutable revision. The returned
// document is a deep copy and can be modified by the caller.
func (s *AnalysisSnapshot) ControlFlowGraphs(load func() (vbacfg.Document, error)) (vbacfg.Document, bool, error) {
	return s.ControlFlowGraphsContext(context.Background(), func(context.Context) (vbacfg.Document, error) { return load() })
}

// ControlFlowGraphsContext has the same retryable cancellation contract as ProcedureIRContext.
func (s *AnalysisSnapshot) ControlFlowGraphsContext(ctx context.Context, load func(context.Context) (vbacfg.Document, error)) (vbacfg.Document, bool, error) {
	if s == nil {
		result, err := load(ctx)
		return vbacfg.CloneDocument(result), false, err
	}
	for {
		if err := ctx.Err(); err != nil {
			return vbacfg.Document{}, false, err
		}
		s.controlFlowMu.Lock()
		if s.controlFlowDone {
			result, err := vbacfg.CloneDocument(s.controlFlow), s.controlFlowErr
			s.controlFlowMu.Unlock()
			return result, true, err
		}
		if s.controlFlowWait != nil {
			wait := s.controlFlowWait
			s.controlFlowMu.Unlock()
			select {
			case <-ctx.Done():
				return vbacfg.Document{}, false, ctx.Err()
			case <-wait:
			}
			continue
		}
		wait := make(chan struct{})
		s.controlFlowWait = wait
		s.controlFlowMu.Unlock()
		defer func() {
			s.controlFlowMu.Lock()
			if s.controlFlowWait == wait {
				s.controlFlowWait = nil
				close(wait)
			}
			s.controlFlowMu.Unlock()
		}()
		result, err := load(ctx)
		if err == nil {
			err = ctx.Err()
		}
		s.controlFlowMu.Lock()
		if !isRetryableContextError(err) {
			s.controlFlow, s.controlFlowErr, s.controlFlowDone = vbacfg.CloneDocument(result), err, true
		}
		s.controlFlowMu.Unlock()
		return vbacfg.CloneDocument(result), false, err
	}
}

func isRetryableContextError(err error) bool {
	return errors.Is(err, context.Canceled) || errors.Is(err, context.DeadlineExceeded)
}

// documentIndex returns the immutable lookup index for this snapshot. The
// caller supplies the same symbol loader used by document analysis so custom
// analyzers and the normal source-symbol cache retain identical semantics.
func (s *AnalysisSnapshot) documentIndex(load DocumentSymbolLoader) (*documentIndex, bool, error) {
	if s == nil {
		return nil, false, errAnalysisSnapshotRetired
	}
	initialized := true
	s.indexOnce.Do(func() {
		initialized = false
		s.initProcedures()
		_, _, err := s.SourceSymbols(load)
		if err != nil {
			s.indexErr = err
			return
		}
		// SourceSymbols returns a defensive copy to public callers. The snapshot
		// owns s.symbols and never mutates it after the successful build completes, so the
		// internal index can safely avoid a second full clone.
		s.index = buildDocumentIndex(s.source, s.lines, s.procedures, s.procedureLines, s.symbols)
	})
	return s.index, initialized, s.indexErr
}

func (s *AnalysisSnapshot) identifiers() [][]byteSpan {
	if s == nil {
		return nil
	}
	s.semanticOnce.Do(func() {
		s.semanticIdentifiers = make([][]byteSpan, len(s.lines))
		for lineNo, line := range s.lines {
			s.semanticIdentifiers[lineNo] = codeIdentifierSpans(line)
		}
	})
	return s.semanticIdentifiers
}

// ParsedDocument returns the snapshot-owned tree-sitter state, creating it at
// most once for this document revision. Callers must not close the returned
// document; the snapshot owns its lifecycle and retires it exactly once.
func (s *AnalysisSnapshot) ParsedDocument() (*ast.ParsedDocument, error) {
	if s == nil {
		return nil, errAnalysisSnapshotRetired
	}
	s.parsedMu.Lock()
	defer s.parsedMu.Unlock()
	if s.parsedDocument != nil || s.parsedErr != nil {
		return s.parsedDocument, s.parsedErr
	}
	if s.retired.Load() {
		return nil, errAnalysisSnapshotRetired
	}
	parse := s.parseDocument
	if parse == nil {
		parse = ast.ParseDocument
	}
	s.parsedDocument, s.parsedErr = parse(s.path, []byte(s.source))
	if s.parsedDocument != nil {
		s.parseCount.Add(1)
	}
	return s.parsedDocument, s.parsedErr
}

// ParseCount reports how many document-owned parses this snapshot created.
// It exists for diagnostics benchmarking and regression tests.
func (s *AnalysisSnapshot) ParseCount() uint64 {
	if s == nil {
		return 0
	}
	return s.parseCount.Load()
}

// FullHashCount reports how many revision comparisons required hashing an
// independently backed source string. It exists for diagnostics benchmarking
// and regression tests; creating the snapshot's canonical digest is excluded.
func (s *AnalysisSnapshot) FullHashCount() uint64 {
	if s == nil {
		return 0
	}
	return s.fullHashCount.Load()
}

// Retire marks the snapshot as no longer owned by its publisher.
// It is idempotent and is the cleanup boundary for future owned resources.
func (s *AnalysisSnapshot) Retire() {
	if s == nil || !s.retired.CompareAndSwap(false, true) {
		return
	}
	s.parsedMu.Lock()
	if s.parsedDocument != nil {
		s.parsedDocument.Close()
	}
	s.parsedMu.Unlock()
}

func (s *AnalysisSnapshot) Retired() bool { return s != nil && s.retired.Load() }

func analysisSnapshotForDocument(doc Document) *AnalysisSnapshot {
	if doc.Snapshot != nil && doc.Snapshot.sameRevision(doc) {
		return doc.Snapshot
	}
	return nil
}

func parsedDocumentForDocument(doc Document) (*ast.ParsedDocument, func(), error) {
	if snapshot := analysisSnapshotForDocument(doc); snapshot != nil {
		parsed, err := snapshot.ParsedDocument()
		return parsed, func() {}, err
	}
	parsed, err := ast.ParseDocument(doc.Path, []byte(doc.Source))
	if err != nil {
		return nil, func() {}, err
	}
	return parsed, parsed.Close, nil
}

func procedureIRForDocumentContext(ctx context.Context, doc Document, rootDir string, parsed *ast.ParsedDocument) (procedureir.DocumentIR, error) {
	load := func(loadCtx context.Context) (procedureir.DocumentIR, error) {
		if snapshot := analysisSnapshotForDocument(doc); snapshot != nil {
			if result, reused, err := snapshot.incrementalProcedureIR(loadCtx, rootDir); err != nil || reused {
				return result, err
			}
		}
		return procedureir.BuildParsedContext(loadCtx, procedureir.BuildOptions{
			RootDir:    rootDir,
			Path:       doc.Path,
			ModuleKind: doc.ModuleKind,
		}, parsed)
	}
	if snapshot := analysisSnapshotForDocument(doc); snapshot != nil {
		result, _, err := snapshot.ProcedureIRContext(ctx, load)
		return result, err
	}
	return load(ctx)
}

func controlFlowForDocumentContext(ctx context.Context, doc Document, ir procedureir.DocumentIR) (vbacfg.Document, error) {
	load := func(loadCtx context.Context) (vbacfg.Document, error) {
		if snapshot := analysisSnapshotForDocument(doc); snapshot != nil {
			if result, reused, err := snapshot.incrementalCFG(loadCtx, ir); err != nil || reused {
				return result, err
			}
		}
		return vbacfg.BuildDocumentContext(loadCtx, ir)
	}
	if snapshot := analysisSnapshotForDocument(doc); snapshot != nil {
		result, _, err := snapshot.ControlFlowGraphsContext(ctx, load)
		if err == nil {
			snapshot.seedCFGArtifacts(ir, result)
		}
		return result, err
	}
	return load(ctx)
}

func documentLines(doc Document) []string {
	if snapshot := analysisSnapshotForDocument(doc); snapshot != nil {
		return snapshot.sourceLines()
	}
	return normalizedLines(doc.Source)
}

func cloneAnalysisSymbols(syms []Symbol) []Symbol {
	out := make([]Symbol, len(syms))
	for i, sym := range syms {
		out[i] = sym
		out[i].Parameters = append([]Parameter(nil), sym.Parameters...)
		out[i].Documentation.ParameterEntries = append(out[i].Documentation.ParameterEntries[:0:0], sym.Documentation.ParameterEntries...)
		if sym.Documentation.Parameters != nil {
			out[i].Documentation.Parameters = make(map[string]string, len(sym.Documentation.Parameters))
			for key, value := range sym.Documentation.Parameters {
				out[i].Documentation.Parameters[key] = value
			}
		}
		if sym.Documentation.UnknownSections != nil {
			out[i].Documentation.UnknownSections = make(map[string]string, len(sym.Documentation.UnknownSections))
			for key, value := range sym.Documentation.UnknownSections {
				out[i].Documentation.UnknownSections[key] = value
			}
		}
	}
	return out
}
