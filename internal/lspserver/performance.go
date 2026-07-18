package lspserver

import (
	"strings"
	"time"

	"github.com/harumiWeb/xlflow/internal/vba/intel"
)

type performanceMeasurement struct {
	server    *Server
	operation string
	document  intel.Document
	started   time.Time
}

func (s *Server) startPerformance(operation string, doc intel.Document) *performanceMeasurement {
	measurement := s.startPerformanceURI(operation, doc.URI)
	measurement.setDocument(doc)
	return measurement
}

func (s *Server) startPerformanceURI(operation, uri string) *performanceMeasurement {
	if !s.opts.PerformanceLog {
		return nil
	}
	return &performanceMeasurement{
		server:    s,
		operation: operation,
		document:  intel.Document{URI: uri},
		started:   time.Now(),
	}
}

func (m *performanceMeasurement) setDocument(doc intel.Document) {
	if m != nil {
		m.document = doc
	}
}

func (m *performanceMeasurement) finish(resultCount int, err error) {
	if m == nil {
		return
	}
	outcome := "ok"
	if err != nil {
		outcome = "error"
	}
	doc := m.document
	m.server.logger.Printf(
		"performance operation=%q uri=%q path=%q version=%d bytes=%d lines=%d elapsed_ms=%.3f result_count=%d outcome=%q",
		m.operation,
		doc.URI,
		doc.Path,
		doc.Version,
		len(doc.Source),
		sourceLineCount(doc.Source),
		float64(time.Since(m.started))/float64(time.Millisecond),
		resultCount,
		outcome,
	)
}

func (m *performanceMeasurement) finishDiagnostics(resultCount int, generation uint64, discarded bool) {
	if m == nil {
		return
	}
	doc := m.document
	outcome := "ok"
	if discarded {
		outcome = "discarded"
	}
	m.server.logger.Printf(
		"performance operation=%q uri=%q path=%q version=%d generation=%d bytes=%d lines=%d elapsed_ms=%.3f result_count=%d outcome=%q discarded=%t",
		m.operation,
		doc.URI,
		doc.Path,
		doc.Version,
		generation,
		len(doc.Source),
		sourceLineCount(doc.Source),
		float64(time.Since(m.started))/float64(time.Millisecond),
		resultCount,
		outcome,
		discarded,
	)
}

func (s *Server) logInitialWorkspaceIndexPerformance(fileCount int, started time.Time, err error) {
	if !s.opts.PerformanceLog {
		return
	}
	outcome := "ok"
	if err != nil {
		outcome = "error"
	}
	s.logger.Printf(
		"performance operation=%q elapsed_ms=%.3f file_count=%d outcome=%q",
		"workspaceSymbols/index/initial",
		float64(time.Since(started))/float64(time.Millisecond),
		fileCount,
		outcome,
	)
}

func (s *Server) logDocumentCachePerformance(operation, cache string, doc intel.Document, resultCount int, started time.Time, err error) {
	if !s.opts.PerformanceLog {
		return
	}
	outcome := "ok"
	if err != nil {
		outcome = "error"
	}
	s.logger.Printf(
		"performance operation=%q uri=%q path=%q version=%d bytes=%d lines=%d elapsed_ms=%.3f result_count=%d outcome=%q cache=%q",
		operation,
		doc.URI,
		doc.Path,
		doc.Version,
		len(doc.Source),
		sourceLineCount(doc.Source),
		float64(time.Since(started))/float64(time.Millisecond),
		resultCount,
		outcome,
		cache,
	)
}

func (s *Server) logDocumentChangePerformance(uri string, version int32, change documentChangeResult, started time.Time) {
	if !s.opts.PerformanceLog {
		return
	}
	doc := change.document
	if doc.URI == "" {
		doc.URI = uri
	}
	// A retained change reports the attempted LSP version, even though the
	// source fields intentionally describe the last valid snapshot.
	doc.Version = version
	s.logger.Printf(
		"performance operation=%q uri=%q path=%q version=%d bytes=%d lines=%d elapsed_ms=%.3f result_count=%d outcome=%q parse_mode=%q fallback_reason=%q",
		"textDocument/didChange/parse",
		doc.URI,
		doc.Path,
		doc.Version,
		len(doc.Source),
		sourceLineCount(doc.Source),
		float64(time.Since(started))/float64(time.Millisecond),
		0,
		map[bool]string{true: "ok", false: "retained"}[change.applied],
		change.parseMode,
		change.fallbackReason,
	)
}

func sourceLineCount(source string) int {
	if source == "" {
		return 0
	}
	return strings.Count(source, "\n") + 1
}

func cacheStatus(hit bool) string {
	if hit {
		return "hit"
	}
	return "miss"
}
