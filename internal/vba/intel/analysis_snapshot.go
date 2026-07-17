package intel

import (
	"crypto/sha256"
	"encoding/hex"
	"strings"
	"sync"
	"sync/atomic"
)

// ProcedureInfo describes the source range occupied by a VBA procedure.
type ProcedureInfo struct {
	Name  string
	Range Range
}

// AnalysisSnapshot is the immutable source state for one document revision.
// Derived values are initialized once and may be read concurrently.
type AnalysisSnapshot struct {
	uri        string
	path       string
	version    int32
	moduleKind string
	source     string
	sourceHash [sha256.Size]byte
	lines      []string

	proceduresOnce sync.Once
	procedures     []ProcedureInfo
	procedureLines []int

	symbolsOnce sync.Once
	symbols     []Symbol
	symbolsErr  error

	semanticOnce        sync.Once
	semanticIdentifiers [][]byteSpan

	retired atomic.Bool
}

// NewAnalysisSnapshot captures doc as an immutable analysis revision.
func NewAnalysisSnapshot(doc Document) *AnalysisSnapshot {
	source := doc.Source
	normalized := strings.ReplaceAll(source, "\r\n", "\n")
	normalized = strings.ReplaceAll(normalized, "\r", "\n")
	return &AnalysisSnapshot{
		uri:        doc.URI,
		path:       doc.Path,
		version:    doc.Version,
		moduleKind: doc.ModuleKind,
		source:     source,
		sourceHash: sha256.Sum256([]byte(source)),
		lines:      strings.Split(normalized, "\n"),
	}
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
	return s != nil && s.uri == doc.URI && s.path == doc.Path &&
		s.version == doc.Version && s.moduleKind == doc.ModuleKind &&
		s.sourceHash == sha256.Sum256([]byte(doc.Source))
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
	if s == nil {
		syms, err := load()
		return cloneAnalysisSymbols(syms), false, err
	}
	initialized := true
	s.symbolsOnce.Do(func() {
		initialized = false
		syms, err := load()
		s.symbols = cloneAnalysisSymbols(syms)
		s.symbolsErr = err
	})
	return cloneAnalysisSymbols(s.symbols), initialized, s.symbolsErr
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

// Retire marks the snapshot as no longer owned by its publisher.
// It is idempotent and is the cleanup boundary for future owned resources.
func (s *AnalysisSnapshot) Retire() {
	if s != nil {
		s.retired.Store(true)
	}
}

func (s *AnalysisSnapshot) Retired() bool { return s != nil && s.retired.Load() }

func analysisSnapshotForDocument(doc Document) *AnalysisSnapshot {
	if doc.Snapshot != nil && doc.Snapshot.sameRevision(doc) {
		return doc.Snapshot
	}
	return nil
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
