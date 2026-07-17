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

func (s *Server) logCachePerformance(operation, cache string, resultCount int, started time.Time, err error) {
	if !s.opts.PerformanceLog {
		return
	}
	outcome := "ok"
	if err != nil {
		outcome = "error"
	}
	s.logger.Printf(
		"performance operation=%q elapsed_ms=%.3f result_count=%d outcome=%q cache=%q",
		operation,
		float64(time.Since(started))/float64(time.Millisecond),
		resultCount,
		outcome,
		cache,
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
