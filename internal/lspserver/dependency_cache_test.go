package lspserver

import (
	"context"
	"errors"
	"io"
	"log"
	"path/filepath"
	"runtime"
	"strings"
	"sync"
	"sync/atomic"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/intel"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func TestDependencyCacheReusesAcrossKeysThatRemainStable(t *testing.T) {
	var cache dependencyCache[int]
	var builds atomic.Int32
	build := func(value int) func() (int, error) {
		return func() (int, error) {
			builds.Add(1)
			return value, nil
		}
	}
	if value, err, hit := cache.getOrBuildContext(context.Background(), "body:unchanged", build(7)); err != nil || hit || value != 7 {
		t.Fatalf("first build = (%d, %v, hit=%v)", value, err, hit)
	}
	if value, err, hit := cache.getOrBuildContext(context.Background(), "body:unchanged", build(9)); err != nil || !hit || value != 7 {
		t.Fatalf("stable dependency lookup = (%d, %v, hit=%v)", value, err, hit)
	}
	if value, err, hit := cache.getOrBuildContext(context.Background(), "body:changed", build(11)); err != nil || hit || value != 11 {
		t.Fatalf("changed dependency lookup = (%d, %v, hit=%v)", value, err, hit)
	}
	if got := builds.Load(); got != 2 {
		t.Fatalf("build count = %d, want 2", got)
	}
}

func TestDependencyCacheCanceledBuildIsRetryable(t *testing.T) {
	var cache dependencyCache[int]
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	called := false
	if _, err, _ := cache.getOrBuildContext(ctx, "cancelled", func() (int, error) {
		called = true
		return 1, nil
	}); !errors.Is(err, context.Canceled) {
		t.Fatalf("canceled build error = %v, want context.Canceled", err)
	}
	if called {
		t.Fatal("canceled build was invoked")
	}
	if value, err, hit := cache.getOrBuildContext(context.Background(), "cancelled", func() (int, error) {
		return 2, nil
	}); err != nil || hit || value != 2 {
		t.Fatalf("retry = (%d, %v, hit=%v)", value, err, hit)
	}
}

func TestDependencyCachePanicIsNotPublished(t *testing.T) {
	var cache dependencyCache[int]
	func() {
		defer func() {
			if recover() == nil {
				t.Fatal("panic from cache builder was not propagated")
			}
		}()
		_, _, _ = cache.getOrBuildContext(context.Background(), "panic", func() (int, error) {
			panic("builder failed")
		})
	}()
	if _, ok := cache.get("panic"); ok {
		t.Fatal("panic result was published")
	}
	if value, err, hit := cache.getOrBuildContext(context.Background(), "panic", func() (int, error) {
		return 9, nil
	}); err != nil || hit || value != 9 {
		t.Fatalf("retry after panic = (%d, %v, hit=%v)", value, err, hit)
	}
}

func TestDependencyCacheSharesOneBuildPerKey(t *testing.T) {
	var cache dependencyCache[int]
	var builds atomic.Int32
	started := make(chan struct{})
	release := make(chan struct{})
	build := func() (int, error) {
		if builds.Add(1) == 1 {
			close(started)
		}
		<-release
		return 42, nil
	}
	results := make([]int, 2)
	errs := make([]error, 2)
	hits := make([]bool, 2)
	var group sync.WaitGroup
	group.Add(2)
	for index := range results {
		go func(index int) {
			defer group.Done()
			results[index], errs[index], hits[index] = cache.getOrBuildContext(context.Background(), "same", build)
		}(index)
	}
	<-started
	close(release)
	group.Wait()
	if got := builds.Load(); got != 1 {
		t.Fatalf("build count = %d, want 1", got)
	}
	for index := range results {
		if errs[index] != nil || results[index] != 42 {
			t.Fatalf("result[%d] = (%d, %v)", index, results[index], errs[index])
		}
	}
	if hits[0] == hits[1] {
		t.Fatalf("single-flight callers should have one builder and one waiter hit: %v", hits)
	}
}

func TestProjectResolutionReusesResolverAndUnchangedDocumentViews(t *testing.T) {
	first := projectTestSnapshot(
		projectTestProcedure("one.bas", "One.Run", "", "", 1, "Run"),
		projectTestProcedure("two.bas", "Two.Run", "", "", 1, "Run"),
	)
	first.Documents[0].Source = "Public Sub Run()\nEnd Sub\n"
	first.Documents[1].Source = "Public Sub Run()\nEnd Sub\n"
	second := first
	second.Revision = first.Revision + 1
	second.Documents = append([]intel.ProjectAnalysisDocument(nil), first.Documents...)
	second.Documents[0].Source = "Public Sub Run()\n    Debug.Print 1\nEnd Sub\n"

	s := &Server{performance: newPerformanceRecorder(true, log.New(io.Discard, "", 0))}
	if _, _, _, err := s.projectResolution(context.Background(), first, true); err != nil {
		t.Fatal(err)
	}
	if _, _, _, err := s.projectResolution(context.Background(), second, true); err != nil {
		t.Fatal(err)
	}
	if got := s.performance.counterTotal(performanceCounterResolutionResolverBuilds); got != 1 {
		t.Fatalf("resolver builds = %d, want 1", got)
	}
	if got := s.performance.counterTotal(performanceCounterResolutionOverlayBuilds); got != 6 {
		t.Fatalf("overlay builds = %d, want 6 (two initial documents plus one edited document)", got)
	}
	if got := s.performance.counterTotal(performanceCounterProjectCacheReusedEntries); got < 1 {
		t.Fatalf("reused resolution entries = %d, want at least 1", got)
	}
}

func TestDependencyFingerprintsSeparateBodyAndDeclarationInputs(t *testing.T) {
	first := projectTestSnapshot(projectTestProcedure("main.bas", "Main.Run", "", "", 1, "Run"))
	first.Documents[0].Source = "Public Sub Run()\nEnd Sub\n"
	second := first
	second.Documents = append([]intel.ProjectAnalysisDocument(nil), first.Documents...)
	second.Documents[0].Source = "Public Sub Run()\n    Debug.Print 1\nEnd Sub\n"
	if projectResolutionFingerprint(first, true, nil) != projectResolutionFingerprint(second, true, nil) {
		t.Fatal("procedure body edit changed the declaration resolver fingerprint")
	}
	if projectResolutionDocumentFingerprint(first.Documents[0], "resolver", true) == projectResolutionDocumentFingerprint(second.Documents[0], "resolver", true) {
		t.Fatal("procedure body edit did not change the document overlay fingerprint")
	}
	second.Documents[0].Version = "1"
	first.Documents[0].Version = "1"
	if projectDocumentContentFingerprint(first.Documents[0]) == projectDocumentContentFingerprint(second.Documents[0]) {
		t.Fatal("recycled editor version hid a changed document body")
	}
	third := first
	third.Documents = append([]intel.ProjectAnalysisDocument(nil), first.Documents...)
	ir := third.Documents[0].IR
	ir.Procedures = append([]procedureir.ProcedureIR(nil), ir.Procedures...)
	ir.Procedures[0].Symbol.Parameters = []procedureir.Parameter{{Name: "value", Type: "Long", Passing: "ByVal"}}
	third.Documents[0].IR = ir
	if projectResolutionFingerprint(first, true, nil) == projectResolutionFingerprint(third, true, nil) {
		t.Fatal("procedure signature edit did not change the resolver fingerprint")
	}
}

func TestSourceLessIRFingerprintIncludesCompleteSemanticPayload(t *testing.T) {
	first := projectTestSnapshot(projectTestProcedure("main.bas", "Main.Run", "", "", 1, "Run"))
	first.Documents[0].Source = ""
	first.Documents[0].Version = ""
	first.Documents[0].ProcedureCatalog = intel.ProcedureCatalog{}
	second := first
	second.Documents = append([]intel.ProjectAnalysisDocument(nil), first.Documents...)
	ir := second.Documents[0].IR
	ir.Procedures = append([]procedureir.ProcedureIR(nil), ir.Procedures...)
	ir.Procedures[0].Expressions = []procedureir.Expression{{ID: 2, Kind: procedureir.ExpressionIdentifier, Text: "changed"}}
	ir.Procedures[0].Accesses = []procedureir.VariableAccess{{ID: 3, Name: "value", Mode: procedureir.AccessRead, Scope: procedureir.ScopeProject}}
	ir.Procedures[0].Calls = []procedureir.CallSite{{ID: 4, Callee: procedureir.Callee{Text: "Target", BaseName: "Target"}, Resolution: procedureir.CallResolution{Status: procedureir.ResolutionUnresolved}}}
	ir.Procedures[0].Statements[0].Control = &procedureir.ControlFlowMetadata{Transfer: procedureir.TransferGoto, Target: "changed"}
	second.Documents[0].IR = ir
	if projectDocumentContentFingerprint(first.Documents[0]) == projectDocumentContentFingerprint(second.Documents[0]) {
		t.Fatal("source-less IR semantic changes did not change the content fingerprint")
	}
}

func TestProjectConstantsFingerprintIgnoresProcedureBody(t *testing.T) {
	first := projectTestSnapshot(projectTestProcedure("main.bas", "Main.Run", "", "", 1, "Run"))
	first.Documents[0].Source = "Public Const Flag = 1\nPublic Sub Run()\nEnd Sub\n"
	first.Documents[0].ProcedureCatalog.Entries[0].StartByte = len("Public Const Flag = 1\n")
	second := first
	second.Documents = append([]intel.ProjectAnalysisDocument(nil), first.Documents...)
	second.Documents[0].Source = "Public Const Flag = 1\nPublic Sub Run()\n    Debug.Print 1\nEnd Sub\n"
	if projectConstantsDependencyFingerprint(first, true, nil) != projectConstantsDependencyFingerprint(second, true, nil) {
		t.Fatal("procedure body edit changed the project constants fingerprint")
	}
	third := second
	third.Documents = append([]intel.ProjectAnalysisDocument(nil), second.Documents...)
	third.Documents[0].Source = "Public Const Flag = 2\nPublic Sub Run()\n    Debug.Print 1\nEnd Sub\n"
	if projectConstantsDependencyFingerprint(first, true, nil) == projectConstantsDependencyFingerprint(third, true, nil) {
		t.Fatal("module preamble edit did not change the project constants fingerprint")
	}
	s := &Server{performance: newPerformanceRecorder(true, log.New(io.Discard, "", 0))}
	s.projectConstants(first, true, nil)
	s.projectConstants(second, true, nil)
	if got := s.performance.counterTotal(performanceCounterProjectCacheHits); got != 1 {
		t.Fatalf("constant cache hits = %d, want 1", got)
	}
}

func TestProjectEffectsCacheRetainsDependencyProductsAcrossRevisions(t *testing.T) {
	path := "main.bas"
	first := projectTestSnapshot(projectTestProcedure(path, "Main.Run", "", "", 1, "Run"))
	second := first
	second.Documents = append([]intel.ProjectAnalysisDocument(nil), first.Documents...)
	second.Documents[0].Source = "Public Sub Run()\n    Debug.Print 1\nEnd Sub\n"
	s := &Server{performance: newPerformanceRecorder(true, log.New(io.Discard, "", 0))}
	s.projectEffectSummaryWithResolution(context.Background(), first, nil, true)
	s.projectEffectSummaryWithResolution(context.Background(), second, nil, true)
	before := s.performance.counterTotal(performanceCounterProjectCacheHits)
	s.projectEffectSummaryWithResolution(context.Background(), first, nil, true)
	if got := s.performance.counterTotal(performanceCounterProjectCacheHits); got <= before {
		t.Fatalf("historical effects product was not reused: before=%d after=%d", before, got)
	}
}

func TestProjectEffectsSkipBuildWhenContextAlreadyCanceled(t *testing.T) {
	project := projectTestSnapshot(projectTestProcedure("main.bas", "Main.Run", "", "", 1, "Run"))
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	s := &Server{performance: newPerformanceRecorder(true, log.New(io.Discard, "", 0))}
	if got := s.projectEffectSummaryWithResolution(ctx, project, nil, true); got.ProcedureCount() != 0 {
		t.Fatalf("canceled effects request returned %d procedures", got.ProcedureCount())
	}
	s.projectEffectsState.mu.Lock()
	building, valid := s.projectEffectsState.building, s.projectEffectsState.valid
	s.projectEffectsState.mu.Unlock()
	if building || valid {
		t.Fatalf("canceled effects request changed state: building=%v valid=%v", building, valid)
	}
	if got := len(s.projectEffectsCache.values); got != 0 {
		t.Fatalf("canceled effects request published %d cache entries", got)
	}
}

func TestProjectEffectsHistoricalCacheRebindsDisplayPaths(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows path casing is the behavior under test")
	}
	root := t.TempDir()
	oldPath := strings.ToUpper(filepath.Join(root, "Module.bas"))
	newPath := strings.ToLower(oldPath)
	first := projectTestSnapshot(projectTestProcedure(oldPath, "Module.Run", "", "", 1, "Run"))
	second := first
	second.Documents = append([]intel.ProjectAnalysisDocument(nil), first.Documents...)
	second.Documents[0].IR.Path = newPath
	s := &Server{performance: newPerformanceRecorder(true, log.New(io.Discard, "", 0))}
	s.projectEffectSummaryWithResolution(context.Background(), first, nil, true)
	got := s.projectEffectSummaryWithResolution(context.Background(), second, nil, true)
	if got.ProcedureCount() != 1 {
		t.Fatalf("historical effects cache returned %d procedures", got.ProcedureCount())
	}
	if display := got.All()[0].Identity.File; display != filepath.ToSlash(newPath) {
		t.Fatalf("historical cache display path = %q, want current source spelling %q", display, filepath.ToSlash(newPath))
	}
}
