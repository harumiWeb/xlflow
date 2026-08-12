package lspserver

import (
	"context"
	"errors"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"testing"
	"time"

	"github.com/tliron/glsp"
	protocol "github.com/tliron/glsp/protocol_3_16"

	"github.com/harumiWeb/xlflow/internal/config"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/callgraph"
	"github.com/harumiWeb/xlflow/internal/vba/calls"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
)

func TestWorkspaceAnalysisIndexParsesOnceAndUpdatesOnlyChangedFile(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	a := filepath.Join(moduleDir, "A.bas")
	b := filepath.Join(moduleDir, "B.bas")
	if err := os.WriteFile(a, []byte("Sub Alpha()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(b, []byte("Sub Beta()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}

	counts := map[string]int{}
	parse := func(_ context.Context, file symbols.SourceFile, source []byte) (indexedFileAnalysis, error) {
		counts[file.Path]++
		name := strings.Fields(string(source))[1]
		name = strings.TrimSuffix(name, "()")
		return indexedFileAnalysis{path: file.Path, moduleKind: file.ModuleKind, symbols: []intel.Symbol{{
			Name: name, Kind: "sub", Module: strings.TrimSuffix(filepath.Base(file.Path), filepath.Ext(file.Path)), File: filepath.ToSlash(file.Path),
		}}}, nil
	}
	index := newWorkspaceAnalysisIndex(root, config.Default(), parse, nil)
	if err := index.waitReady(); err != nil {
		t.Fatal(err)
	}
	if counts[a] != 1 || counts[b] != 1 {
		t.Fatalf("initial parses = %#v, want one per file", counts)
	}
	if got, err := index.searchExact("Alpha"); err != nil || len(got) != 1 {
		t.Fatalf("exact = %#v, %v", got, err)
	}
	if got, err := index.searchPrefix("Be"); err != nil || len(got) != 1 || got[0].Name != "Beta" {
		t.Fatalf("prefix = %#v, %v", got, err)
	}
	if got, err := index.searchQualified("A.Alpha"); err != nil || len(got) != 1 {
		t.Fatalf("qualified = %#v, %v", got, err)
	}
	if got, err := index.searchModule("B"); err != nil || len(got) != 1 {
		t.Fatalf("module = %#v, %v", got, err)
	}
	if got, err := index.searchKind("sub"); err != nil || len(got) != 2 {
		t.Fatalf("kind = %#v, %v", got, err)
	}
	if got, err := index.searchContains("alp"); err != nil || len(got) != 1 {
		t.Fatalf("contains = %#v, %v", got, err)
	}
	if counts[a] != 1 || counts[b] != 1 {
		t.Fatalf("warm queries reparsed files: %#v", counts)
	}

	if err := os.WriteFile(a, []byte("Sub Updated()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := index.updatePath(a); err != nil {
		t.Fatal(err)
	}
	if got, _ := index.searchExact("Alpha"); len(got) != 0 {
		t.Fatalf("old symbol remains: %#v", got)
	}
	if got, _ := index.searchExact("Updated"); len(got) != 1 {
		t.Fatalf("updated symbol missing: %#v", got)
	}
	if counts[a] != 2 || counts[b] != 1 {
		t.Fatalf("changed file parses = %#v", counts)
	}
	if err := index.updatePath(a); err != nil {
		t.Fatal(err)
	}
	if counts[a] != 2 || counts[b] != 1 {
		t.Fatalf("duplicate watcher event reparsed: %#v", counts)
	}
}

func TestWorkspaceAnalysisIndexParseFailureMarksSnapshotIncomplete(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "src", "modules", "Main.bas")
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte("good"), 0o644); err != nil {
		t.Fatal(err)
	}
	parse := func(_ context.Context, file symbols.SourceFile, source []byte) (indexedFileAnalysis, error) {
		if strings.TrimSpace(string(source)) == "bad" {
			return indexedFileAnalysis{}, errors.New("parse failed")
		}
		return indexedFileAnalysis{
			path: file.Path, moduleKind: file.ModuleKind,
			symbols: []intel.Symbol{{Name: "Good", File: file.Path}},
		}, nil
	}
	index := newWorkspaceAnalysisIndex(root, config.Default(), parse, nil)
	if err := index.waitReady(); err != nil {
		t.Fatal(err)
	}
	if snapshot := index.projectSnapshot(); !snapshot.Complete {
		t.Fatalf("initial snapshot = %#v, want complete", snapshot)
	}

	if err := os.WriteFile(path, []byte("bad"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := index.updatePath(path); err == nil {
		t.Fatal("parse failure was not returned")
	}
	if got, err := index.searchExact("Good"); err != nil || len(got) != 1 {
		t.Fatalf("effective symbols were dropped after parse failure: %+v, %v", got, err)
	}
	if snapshot := index.projectSnapshot(); snapshot.Complete {
		t.Fatalf("failed parse published complete snapshot: %#v", snapshot)
	}
	if err := os.Remove(path); err != nil {
		t.Fatal(err)
	}
	if err := index.removePath(path); err != nil {
		t.Fatal(err)
	}
	if snapshot := index.projectSnapshot(); !snapshot.Complete {
		t.Fatalf("known deleted path left snapshot incomplete: %#v", snapshot)
	}

	if err := os.WriteFile(path, []byte("fixed"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := index.updatePath(path); err != nil {
		t.Fatal(err)
	}
	if snapshot := index.projectSnapshot(); !snapshot.Complete {
		t.Fatalf("clean reparse left snapshot incomplete: %#v", snapshot)
	}
}

func TestWorkspaceAnalysisIndexRecoveredIRMarksSnapshotIncomplete(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "src", "modules", "Main.bas")
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte("clean"), 0o644); err != nil {
		t.Fatal(err)
	}
	parse := func(_ context.Context, file symbols.SourceFile, source []byte) (indexedFileAnalysis, error) {
		ir := procedureir.DocumentIR{Path: file.Path, ModuleName: "Main", ModuleKind: file.ModuleKind}
		if strings.TrimSpace(string(source)) == "recovered" {
			ir.Parse.HasMissing = true
		}
		return indexedFileAnalysis{path: file.Path, moduleKind: file.ModuleKind, procedureIR: ir}, nil
	}
	index := newWorkspaceAnalysisIndex(root, config.Default(), parse, nil)
	if err := index.waitReady(); err != nil {
		t.Fatal(err)
	}
	if snapshot := index.projectSnapshot(); !snapshot.Complete {
		t.Fatalf("initial snapshot = %#v, want complete", snapshot)
	}

	if err := os.WriteFile(path, []byte("recovered"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := index.updatePath(path); err != nil {
		t.Fatal(err)
	}
	if snapshot := index.projectSnapshot(); snapshot.Complete {
		t.Fatalf("recovered IR published complete snapshot: %#v", snapshot)
	}

	if err := os.WriteFile(path, []byte("clean-again"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := index.updatePath(path); err != nil {
		t.Fatal(err)
	}
	if snapshot := index.projectSnapshot(); !snapshot.Complete {
		t.Fatalf("clean IR left snapshot incomplete: %#v", snapshot)
	}
}

func TestWorkspaceAnalysisIndexInitialParseFailureCanRecoverToCompleteSnapshot(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "src", "modules", "Main.bas")
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte("bad"), 0o644); err != nil {
		t.Fatal(err)
	}
	parse := func(_ context.Context, file symbols.SourceFile, source []byte) (indexedFileAnalysis, error) {
		if strings.TrimSpace(string(source)) == "bad" {
			return indexedFileAnalysis{}, errors.New("initial parse failed")
		}
		return indexedFileAnalysis{path: file.Path, moduleKind: file.ModuleKind}, nil
	}
	index := newWorkspaceAnalysisIndex(root, config.Default(), parse, nil)
	if err := index.waitReady(); err != nil {
		t.Fatalf("initial discovery failed after one file parse error: %v", err)
	}
	if snapshot := index.projectSnapshot(); snapshot.Complete {
		t.Fatalf("initial failed parse published complete snapshot: %#v", snapshot)
	}

	if err := os.WriteFile(path, []byte("fixed"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := index.updatePath(path); err != nil {
		t.Fatal(err)
	}
	if snapshot := index.projectSnapshot(); !snapshot.Complete {
		t.Fatalf("clean refresh left initial snapshot incomplete: %#v", snapshot)
	}
}

func TestCallGraphSnapshotUsesEffectiveTypeReferenceFacts(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(moduleDir, "Main.bas")
	if err := os.WriteFile(path, []byte("DiskType"), 0o644); err != nil {
		t.Fatal(err)
	}
	parse := func(_ context.Context, file symbols.SourceFile, source []byte) (indexedFileAnalysis, error) {
		name := strings.TrimSpace(string(source))
		return indexedFileAnalysis{
			path: file.Path, moduleKind: file.ModuleKind,
			typeReferences: []calls.TypeReference{{Kind: "uses_type", File: "src/modules/Main.bas", Module: "Main", Target: name}},
		}, nil
	}
	index := newWorkspaceAnalysisIndex(root, config.Default(), parse, nil)
	first, err := index.callGraphSnapshot()
	if err != nil || len(first.TypeReferences) != 1 || first.TypeReferences[0].Target != "DiskType" {
		t.Fatalf("disk snapshot = %+v, %v", first, err)
	}
	doc := intel.Document{Path: path, Source: "OverlayType", ModuleKind: "standard", Version: 1}
	index.setOverlay(doc, indexedFileAnalysis{typeReferences: []calls.TypeReference{{Kind: "constructs", File: "src/modules/Main.bas", Module: "Main", Target: "OverlayType"}}})
	overlay, err := index.callGraphSnapshot()
	if err != nil || len(overlay.TypeReferences) != 1 || overlay.TypeReferences[0].Kind != "constructs" || overlay.TypeReferences[0].Target != "OverlayType" {
		t.Fatalf("overlay snapshot = %+v, %v", overlay, err)
	}
}

func TestCallGraphSnapshotPreservesPrivateProcedureVisibility(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	for _, name := range []string{"A.bas", "B.bas"} {
		if err := os.WriteFile(filepath.Join(moduleDir, name), []byte(name), 0o644); err != nil {
			t.Fatal(err)
		}
	}
	parse := func(_ context.Context, file symbols.SourceFile, _ []byte) (indexedFileAnalysis, error) {
		name := strings.TrimSuffix(filepath.Base(file.Path), filepath.Ext(file.Path))
		entry := indexedFileAnalysis{path: file.Path, moduleKind: file.ModuleKind, symbols: []intel.Symbol{{Name: name, Kind: "module", Module: name, ModuleKind: "standard", File: file.Path}, {Name: map[string]string{"A": "Run", "B": "Work"}[name], Kind: "sub", Module: name, ModuleKind: "standard", File: file.Path, Range: intel.Range{Start: intel.Position{Line: 1}}}}}
		if name == "A" {
			entry.callSites = []calls.CallSite{{File: "src/modules/A.bas", Module: "A", Caller: &calls.Caller{Name: "Run", Kind: "sub", QualifiedName: "A.Run"}, Callee: calls.Callee{Text: "Work", BaseName: "Work"}}}
		} else {
			entry.symbols[1].Visibility = "private"
		}
		return entry, nil
	}
	index := newWorkspaceAnalysisIndex(root, config.Default(), parse, nil)
	snapshot, err := index.callGraphSnapshot()
	if err != nil {
		t.Fatal(err)
	}
	if len(snapshot.Calls) != 1 || snapshot.Calls[0].Resolution.Status != "unresolved" {
		t.Fatalf("private cross-module call was resolved: %+v", snapshot.Calls)
	}
}

func TestWorkspaceAnalysisIndexPendingOverlayMasksAndRejectsStalePublish(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "src", "modules", "Main.bas")
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte("Sub DiskName()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	parse := func(_ context.Context, file symbols.SourceFile, _ []byte) (indexedFileAnalysis, error) {
		return indexedFileAnalysis{path: file.Path, moduleKind: file.ModuleKind, symbols: []intel.Symbol{{Name: "DiskName", File: file.Path}}}, nil
	}
	index := newWorkspaceAnalysisIndex(root, config.Default(), parse, nil)
	if err := index.waitReady(); err != nil {
		t.Fatal(err)
	}
	doc := intel.Document{URI: pathToFileURI(path), Path: path, ModuleKind: "standard", Version: 1, Source: "Sub OpenName()\nEnd Sub\n"}
	index.beginOverlay(doc, 1)
	if got, _ := index.searchExact("DiskName"); len(got) != 0 {
		t.Fatalf("pending overlay exposed disk symbols: %+v", got)
	}
	doc.Version = 2
	doc.Source = "Sub LatestName()\nEnd Sub\n"
	index.beginOverlay(doc, 2)
	stale := indexedFileAnalysis{symbols: []intel.Symbol{{Name: "StaleName", File: path}}}
	if index.publishOverlay(doc, 1, stale) {
		t.Fatal("stale overlay generation was published")
	}
	latest := indexedFileAnalysis{symbols: []intel.Symbol{{Name: "LatestName", File: path}}}
	if !index.publishOverlay(doc, 2, latest) {
		t.Fatal("latest overlay generation was rejected")
	}
	if got, _ := index.searchExact("LatestName"); len(got) != 1 {
		t.Fatalf("latest overlay missing: %+v", got)
	}
	if got, _ := index.searchExact("StaleName"); len(got) != 0 {
		t.Fatalf("stale overlay visible: %+v", got)
	}
}

func TestWorkspaceAnalysisIndexAbandonOverlayClearsReservationAndKeepsDiskMasked(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "src", "modules", "Main.bas")
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte("Sub DiskName()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	parse := func(_ context.Context, file symbols.SourceFile, _ []byte) (indexedFileAnalysis, error) {
		return indexedFileAnalysis{symbols: []intel.Symbol{{Name: "DiskName", File: file.Path}}}, nil
	}
	index := newWorkspaceAnalysisIndex(root, config.Default(), parse, nil)
	if err := index.waitReady(); err != nil {
		t.Fatal(err)
	}
	doc := intel.Document{URI: pathToFileURI(path), Path: path, ModuleKind: "standard", Version: 1, Source: "Sub OpenName()\nEnd Sub\n"}
	index.beginOverlay(doc, 1)
	if !index.abandonOverlay(doc, 1) {
		t.Fatal("current failed overlay was not abandoned")
	}
	key := documentSymbolKey(doc)
	index.mu.RLock()
	pending := index.pending[key]
	_, masked := index.overlays[key]
	index.mu.RUnlock()
	if pending != 0 || !masked {
		t.Fatalf("abandoned state = (pending=%d, masked=%t)", pending, masked)
	}
	if err := index.updatePath(path); err != nil {
		t.Fatal(err)
	}
	if got, err := index.searchExact("DiskName"); err != nil || len(got) != 0 {
		t.Fatalf("watcher exposed saved symbols after terminal overlay failure: %+v, %v", got, err)
	}
}

func TestWorkspaceAnalysisIndexPendingOverlayCancelsCloseRefresh(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "src", "modules", "Main.bas")
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte("Sub DiskName()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	index := newWorkspaceAnalysisIndex(root, config.Default(), func(_ context.Context, file symbols.SourceFile, _ []byte) (indexedFileAnalysis, error) {
		return indexedFileAnalysis{symbols: []intel.Symbol{{Name: "DiskName", File: file.Path}}}, nil
	}, nil)
	if err := index.waitReady(); err != nil {
		t.Fatal(err)
	}
	started := make(chan struct{})
	canceled := make(chan struct{})
	index.parse = func(ctx context.Context, _ symbols.SourceFile, _ []byte) (indexedFileAnalysis, error) {
		close(started)
		<-ctx.Done()
		close(canceled)
		return indexedFileAnalysis{}, ctx.Err()
	}
	refresh := index.beginClearOverlay(path)
	done := make(chan error, 1)
	go func() {
		_, _, err := index.finishClearOverlay(refresh)
		done <- err
	}()
	<-started
	doc := intel.Document{URI: pathToFileURI(path), Path: path, ModuleKind: "standard", Version: 1, Source: "Sub OpenName()\nEnd Sub\n"}
	index.beginOverlay(doc, 1)
	select {
	case <-canceled:
	case <-time.After(time.Second):
		t.Fatal("new overlay did not cancel the close refresh parse")
	}
	if err := <-done; err != nil {
		t.Fatalf("canceled close refresh = %v", err)
	}
}

func TestWorkspaceAnalysisIndexPendingOverlayCancelsInitialDiskParse(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "src", "modules", "Main.bas")
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte("Sub DiskName()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	started := make(chan struct{})
	canceled := make(chan struct{})
	parse := func(ctx context.Context, file symbols.SourceFile, _ []byte) (indexedFileAnalysis, error) {
		close(started)
		<-ctx.Done()
		close(canceled)
		return indexedFileAnalysis{}, ctx.Err()
	}
	index := newWorkspaceAnalysisIndex(root, config.Default(), parse, nil)
	index.start()
	select {
	case <-started:
	case <-time.After(time.Second):
		t.Fatal("initial disk parse did not start")
	}

	doc := intel.Document{URI: pathToFileURI(path), Path: path, ModuleKind: "standard", Version: 1, Source: "Sub OpenName()\nEnd Sub\n"}
	index.beginOverlay(doc, 1)
	select {
	case <-canceled:
	case <-time.After(time.Second):
		t.Fatal("pending overlay did not cancel the initial disk parse")
	}
	if err := index.waitReady(); err != nil {
		t.Fatalf("canceled masked disk parse failed initial scan: %v", err)
	}
	if got, err := index.searchExact("DiskName"); err != nil || len(got) != 0 {
		t.Fatalf("canceled disk analysis leaked through pending overlay: %+v, %v", got, err)
	}
	if !index.publishOverlay(doc, 1, indexedFileAnalysis{symbols: []intel.Symbol{{Name: "OpenName", File: path}}}) {
		t.Fatal("current overlay was rejected after canceling disk parse")
	}
	if got, err := index.searchExact("OpenName"); err != nil || len(got) != 1 {
		t.Fatalf("published overlay missing: %+v, %v", got, err)
	}
}

func TestWorkspaceAnalysisIndexCloseRefreshStaysMaskedUntilFreshDiskPublish(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "src", "modules", "Main.bas")
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte("Sub DiskName()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	parse := func(_ context.Context, file symbols.SourceFile, _ []byte) (indexedFileAnalysis, error) {
		return indexedFileAnalysis{symbols: []intel.Symbol{{Name: "DiskName", File: file.Path}}}, nil
	}
	index := newWorkspaceAnalysisIndex(root, config.Default(), parse, nil)
	if err := index.waitReady(); err != nil {
		t.Fatal(err)
	}
	doc := intel.Document{URI: pathToFileURI(path), Path: path, ModuleKind: "standard", Version: 1, Source: "Sub OpenName()\nEnd Sub\n"}
	index.setOverlay(doc, indexedFileAnalysis{symbols: []intel.Symbol{{Name: "OpenName", File: path}}})

	started := make(chan struct{})
	release := make(chan struct{})
	index.parse = func(_ context.Context, file symbols.SourceFile, _ []byte) (indexedFileAnalysis, error) {
		close(started)
		<-release
		return indexedFileAnalysis{symbols: []intel.Symbol{{Name: "FreshDiskName", File: file.Path}}}, nil
	}
	refreshDone := make(chan error, 1)
	go func() {
		_, _, err := index.clearOverlay(path)
		refreshDone <- err
	}()
	select {
	case <-started:
	case <-time.After(time.Second):
		t.Fatal("close refresh parse did not start")
	}
	for _, name := range []string{"OpenName", "DiskName", "FreshDiskName"} {
		if got, err := index.searchExact(name); err != nil || len(got) != 0 {
			t.Fatalf("%s visible during close refresh: %+v, %v", name, got, err)
		}
	}
	close(release)
	if err := <-refreshDone; err != nil {
		t.Fatal(err)
	}
	if got, err := index.searchExact("FreshDiskName"); err != nil || len(got) != 1 {
		t.Fatalf("fresh disk symbols missing after refresh: %+v, %v", got, err)
	}
}

func TestWorkspaceAnalysisIndexDiskRefreshMarksSnapshotIncompleteWhileParsing(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "src", "modules", "Main.bas")
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte("disk"), 0o644); err != nil {
		t.Fatal(err)
	}
	started := make(chan struct{})
	release := make(chan struct{})
	parse := func(_ context.Context, file symbols.SourceFile, source []byte) (indexedFileAnalysis, error) {
		if strings.TrimSpace(string(source)) == "updated" {
			close(started)
			<-release
		}
		return indexedFileAnalysis{path: file.Path, moduleKind: file.ModuleKind}, nil
	}
	index := newWorkspaceAnalysisIndex(root, config.Default(), parse, nil)
	if err := index.waitReady(); err != nil {
		t.Fatal(err)
	}
	if snapshot := index.projectSnapshot(); !snapshot.Complete {
		t.Fatalf("initial snapshot = %#v, want complete", snapshot)
	}
	if err := os.WriteFile(path, []byte("updated"), 0o644); err != nil {
		t.Fatal(err)
	}
	done := make(chan error, 1)
	go func() { done <- index.updatePath(path) }()
	select {
	case <-started:
	case <-time.After(time.Second):
		t.Fatal("disk refresh parse did not start")
	}
	if snapshot := index.projectSnapshot(); snapshot.Complete {
		t.Fatalf("in-flight disk refresh published complete snapshot: %#v", snapshot)
	}
	close(release)
	if err := <-done; err != nil {
		t.Fatal(err)
	}
	if snapshot := index.projectSnapshot(); !snapshot.Complete {
		t.Fatalf("published disk refresh left snapshot incomplete: %#v", snapshot)
	}
}

func TestServerDiskAnalysisUsesBackgroundAnalysisPermit(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	s.analysisPermits = make(chan struct{}, 1)
	s.analysisPermits <- struct{}{}
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	_, err = s.parseIndexedFileContext(ctx, symbols.SourceFile{
		Path: filepath.Join(root, "src", "modules", "Main.bas"), ModuleKind: "standard",
	}, []byte("Sub Main()\nEnd Sub\n"))
	if !errors.Is(err, context.Canceled) {
		t.Fatalf("disk analysis waiting on full permit = %v, want context.Canceled", err)
	}
	if got := len(s.analysisPermits); got != 1 {
		t.Fatalf("disk analysis changed occupied permit count to %d, want 1", got)
	}
	<-s.analysisPermits
}

func TestDocumentVersionReusesMatchingSnapshotHash(t *testing.T) {
	snapshot := intel.NewAnalysisSnapshot(intel.Document{URI: "file:///Main.bas", Path: "Main.bas", Source: "abc", ModuleKind: "standard", Version: 1})
	defer snapshot.Retire()
	doc := snapshot.Document()
	before := snapshot.FullHashCount()
	if got := documentVersion(doc); got != snapshot.SourceHash() {
		t.Fatalf("document version = %q, want snapshot hash %q", got, snapshot.SourceHash())
	}
	if got := snapshot.FullHashCount() - before; got != 0 {
		t.Fatalf("same-snapshot full hashes = %d, want 0", got)
	}
	stale := doc
	stale.Source = "abd"
	if sameScheduledDocument(doc, stale) {
		t.Fatal("source-mutated snapshot view was treated as the same scheduled document")
	}
}

func TestWorkspaceAnalysisIndexWatcherAndOpenOverlay(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(moduleDir, "Main.bas")
	if err := os.WriteFile(path, []byte("Sub SavedName()\n  SavedTarget\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	if s.handler.WorkspaceDidChangeWatchedFiles == nil {
		t.Fatal("watcher handler was not registered")
	}
	ctx := &glsp.Context{Notify: func(string, any) {}}
	uri := pathToFileURI(path)
	if err := s.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{URI: protocol.DocumentUri(uri), Version: 1, Text: "Sub OpenName()\n  OpenTarget\nEnd Sub\n"}}); err != nil {
		t.Fatal(err)
	}
	if got, _ := s.analysis.searchExact("SavedName"); len(got) != 0 {
		t.Fatalf("pending open leaked saved symbol: %#v", got)
	}
	waitForWorkspaceSymbol(t, s.analysis, "OpenName")
	waitForWorkspaceCall(t, s.analysis, "OpenTarget")
	if err := s.didChange(ctx, &protocol.DidChangeTextDocumentParams{
		TextDocument: protocol.VersionedTextDocumentIdentifier{
			TextDocumentIdentifier: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
			Version:                2,
		},
		ContentChanges: []any{protocol.TextDocumentContentChangeEvent{
			Text: "Sub OpenName()\n  ChangedTarget\nEnd Sub\n",
		}},
	}); err != nil {
		t.Fatal(err)
	}
	if got, _ := s.analysis.searchExact("SavedName"); len(got) != 0 {
		t.Fatalf("pending change leaked saved symbols: %#v", got)
	}
	waitForWorkspaceCall(t, s.analysis, "ChangedTarget")
	if got, err := s.analysis.queryResolvedCalls(workspaceCallQuery{CalleeBase: "OpenTarget"}); err != nil || len(got) != 0 {
		t.Fatalf("superseded open call remains = %+v, %v", got, err)
	}
	if err := os.WriteFile(path, []byte("Sub DiskChanged()\n  DiskTarget\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := s.didChangeWatchedFiles(ctx, &protocol.DidChangeWatchedFilesParams{Changes: []protocol.FileEvent{{URI: protocol.DocumentUri(uri), Type: protocol.FileChangeTypeChanged}}}); err != nil {
		t.Fatal(err)
	}
	if got, _ := s.analysis.searchExact("OpenName"); len(got) != 1 {
		t.Fatalf("disk event replaced open overlay: %#v", got)
	}
	if got, err := s.analysis.queryResolvedCalls(workspaceCallQuery{CalleeBase: "ChangedTarget"}); err != nil || len(got) != 1 {
		t.Fatalf("disk event replaced open call overlay = %+v, %v", got, err)
	}
	if got, err := s.analysis.queryResolvedCalls(workspaceCallQuery{CalleeBase: "DiskTarget"}); err != nil || len(got) != 0 {
		t.Fatalf("hidden disk call leaked through overlay = %+v, %v", got, err)
	}
	if err := s.didClose(ctx, &protocol.DidCloseTextDocumentParams{TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)}}); err != nil {
		t.Fatal(err)
	}
	waitForWorkspaceSymbol(t, s.analysis, "DiskChanged")
	waitForWorkspaceCall(t, s.analysis, "DiskTarget")
	if got, err := s.analysis.queryResolvedCalls(workspaceCallQuery{CalleeBase: "ChangedTarget"}); err != nil || len(got) != 0 {
		t.Fatalf("closed overlay call remains = %+v, %v", got, err)
	}
	if err := os.Remove(path); err != nil {
		t.Fatal(err)
	}
	if err := s.didChangeWatchedFiles(ctx, &protocol.DidChangeWatchedFilesParams{Changes: []protocol.FileEvent{{URI: protocol.DocumentUri(uri), Type: protocol.FileChangeTypeDeleted}}}); err != nil {
		t.Fatal(err)
	}
	if got, _ := s.analysis.searchContains("Name"); len(got) != 0 {
		t.Fatalf("deleted source remains: %#v", got)
	}
	if got, err := s.analysis.queryResolvedCalls(workspaceCallQuery{}); err != nil || len(got) != 0 {
		t.Fatalf("deleted calls remain = %+v, %v", got, err)
	}
}

func TestCachedWorkspaceSymbolQueryUsesCurrentSnapshotWhileOverlayPending(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "src", "modules", "Main.bas")
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte("Sub SavedName()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	s.analysisPermits = make(chan struct{}, 1)
	s.analysisPermits <- struct{}{}
	ctx := &glsp.Context{Notify: func(string, any) {}}
	uri := pathToFileURI(path)
	if err := s.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
		URI: protocol.DocumentUri(uri), Version: 1, Text: "Sub CurrentName()\nEnd Sub\n",
	}}); err != nil {
		t.Fatal(err)
	}
	open := s.docs.openDocuments()
	got, err := s.cachedWorkspaceSymbolQuery(open, intel.WorkspaceSymbolQuery{Text: "CurrentName", Mode: intel.WorkspaceSymbolQueryExact})
	if err != nil || len(got) != 1 || got[0].Name != "CurrentName" {
		t.Fatalf("current snapshot symbols while overlay pending = %+v, %v", got, err)
	}
	if stale, err := s.cachedWorkspaceSymbolQuery(open, intel.WorkspaceSymbolQuery{Text: "SavedName", Mode: intel.WorkspaceSymbolQueryExact}); err != nil || len(stale) != 0 {
		t.Fatalf("saved symbols leaked while overlay pending = %+v, %v", stale, err)
	}
	<-s.analysisPermits
}

func waitForWorkspaceSymbol(t *testing.T, index *workspaceAnalysisIndex, name string) {
	t.Helper()
	deadline := time.Now().Add(5 * time.Second)
	for time.Now().Before(deadline) {
		if got, err := index.searchExact(name); err == nil && len(got) == 1 {
			return
		}
		time.Sleep(10 * time.Millisecond)
	}
	got, err := index.searchExact(name)
	t.Fatalf("workspace symbol %q was not published: %+v, %v", name, got, err)
}

func waitForWorkspaceCall(t *testing.T, index *workspaceAnalysisIndex, name string) {
	t.Helper()
	deadline := time.Now().Add(5 * time.Second)
	for time.Now().Before(deadline) {
		if got, err := index.queryResolvedCalls(workspaceCallQuery{CalleeBase: name}); err == nil && len(got) == 1 {
			return
		}
		time.Sleep(10 * time.Millisecond)
	}
	got, err := index.queryResolvedCalls(workspaceCallQuery{CalleeBase: name})
	t.Fatalf("workspace call %q was not published: %+v, %v", name, got, err)
}

func TestWorkspaceAnalysisIndexUserFormSidecarCallLifecycle(t *testing.T) {
	root := t.TempDir()
	formsDir := filepath.Join(root, "src", "forms")
	codeDir := filepath.Join(formsDir, "code")
	if err := os.MkdirAll(codeDir, 0o755); err != nil {
		t.Fatal(err)
	}
	formPath := filepath.Join(formsDir, "Screen.frm")
	sidecarPath := filepath.Join(codeDir, "Screen.bas")
	if err := os.WriteFile(formPath, []byte("VERSION 5.00\nBegin VB.UserForm Screen\nEnd\nAttribute VB_Name = \"Screen\"\nPrivate Sub FormRun()\n  FrmTarget\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}

	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	if err := s.analysis.waitReady(); err != nil {
		t.Fatal(err)
	}
	if got, err := s.analysis.queryResolvedCalls(workspaceCallQuery{CalleeBase: "FrmTarget"}); err != nil || len(got) != 1 {
		t.Fatalf("frm calls = %+v, %v", got, err)
	}

	if err := os.WriteFile(sidecarPath, []byte("Private Sub FormRun()\n  SidecarTarget\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	ctx := &glsp.Context{Notify: func(string, any) {}}
	sidecarURI := pathToFileURI(sidecarPath)
	if err := s.didChangeWatchedFiles(ctx, &protocol.DidChangeWatchedFilesParams{Changes: []protocol.FileEvent{{
		URI: protocol.DocumentUri(sidecarURI), Type: protocol.FileChangeTypeCreated,
	}}}); err != nil {
		t.Fatal(err)
	}
	if got, err := s.analysis.queryResolvedCalls(workspaceCallQuery{CalleeBase: "SidecarTarget"}); err != nil || len(got) != 1 || got[0].Module != "Screen" {
		t.Fatalf("sidecar calls = %+v, %v", got, err)
	}
	if got, err := s.analysis.queryResolvedCalls(workspaceCallQuery{CalleeBase: "FrmTarget"}); err != nil || len(got) != 0 {
		t.Fatalf("generated frm calls remained in sidecar mode = %+v, %v", got, err)
	}

	if err := os.Remove(sidecarPath); err != nil {
		t.Fatal(err)
	}
	if err := s.didChangeWatchedFiles(ctx, &protocol.DidChangeWatchedFilesParams{Changes: []protocol.FileEvent{{
		URI: protocol.DocumentUri(sidecarURI), Type: protocol.FileChangeTypeDeleted,
	}}}); err != nil {
		t.Fatal(err)
	}
	if got, err := s.analysis.queryResolvedCalls(workspaceCallQuery{CalleeBase: "FrmTarget"}); err != nil || len(got) != 1 {
		t.Fatalf("frm calls not restored after sidecar deletion = %+v, %v", got, err)
	}
	if got, err := s.analysis.queryResolvedCalls(workspaceCallQuery{CalleeBase: "SidecarTarget"}); err != nil || len(got) != 0 {
		t.Fatalf("deleted sidecar calls remain = %+v, %v", got, err)
	}
}

func TestWorkspaceAnalysisIndexIncludesConfiguredRootsAndDuplicateBasenames(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	cfg.Src.Modules = "vba/modules"
	cfg.Src.Classes = "vba/classes"
	cfg.Src.Forms = "vba/forms"
	cfg.Src.Workbook = "vba/workbook"
	files := map[string]string{
		"vba/modules/Shared.bas": "Sub ModuleOnly()\nEnd Sub\n",
		"vba/classes/Shared.cls": "Public Sub ClassOnly()\nEnd Sub\n",
		"vba/workbook/This.bas":  "Sub WorkbookOnly()\nEnd Sub\n",
		"vba/forms/Screen.frm":   "Sub FormOnly()\nEnd Sub\n",
		"outside/Ignored.bas":    "Sub Ignored()\nEnd Sub\n",
	}
	for relative, source := range files {
		path := filepath.Join(root, filepath.FromSlash(relative))
		if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(path, []byte(source), 0o644); err != nil {
			t.Fatal(err)
		}
	}
	s, cleanup, err := New(Options{RootDir: root, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	if err := s.analysis.waitReady(); err != nil {
		t.Fatal(err)
	}
	for _, name := range []string{"ModuleOnly", "ClassOnly", "WorkbookOnly", "FormOnly"} {
		if got, _ := s.analysis.searchExact(name); len(got) != 1 {
			t.Fatalf("configured-root symbol %q = %#v", name, got)
		}
	}
	if got, _ := s.analysis.searchExact("Ignored"); len(got) != 0 {
		t.Fatalf("outside-root symbol indexed: %#v", got)
	}
}

func TestWorkspaceAnalysisIndexExcludesExternalOpenDocumentAndUsesAbsoluteDiscoveryPaths(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	inside := filepath.Join(moduleDir, "Main.bas")
	if err := os.WriteFile(inside, []byte("Sub Inside()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	t.Chdir(root)
	files, err := symbols.DiscoverSourceFiles(symbols.Options{RootDir: ".", Config: config.Default()})
	if err != nil || len(files) != 1 || !filepath.IsAbs(files[0].Path) {
		t.Fatalf("relative-root discovery = %#v, %v", files, err)
	}

	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	if err := s.analysis.waitReady(); err != nil {
		t.Fatal(err)
	}
	external := filepath.Join(t.TempDir(), "External.bas")
	if err := os.WriteFile(external, []byte("Sub ExternalOnly()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	ctx := &glsp.Context{Notify: func(string, any) {}}
	uri := pathToFileURI(external)
	if err := s.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{URI: protocol.DocumentUri(uri), Version: 1, Text: "Sub ExternalOnly()\nEnd Sub\n"}}); err != nil {
		t.Fatal(err)
	}
	if got, err := s.analysis.searchExact("ExternalOnly"); err != nil || len(got) != 0 {
		t.Fatalf("external open document leaked into index: %#v, %v", got, err)
	}
}

func TestWorkspaceAnalysisIndexReresolvesCallsWithoutReparsingCaller(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	a := filepath.Join(moduleDir, "A.bas")
	b := filepath.Join(moduleDir, "B.bas")
	if err := os.WriteFile(a, []byte("caller"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(b, []byte("empty"), 0o644); err != nil {
		t.Fatal(err)
	}

	var mu sync.Mutex
	counts := map[string]int{}
	parse := func(_ context.Context, file symbols.SourceFile, source []byte) (indexedFileAnalysis, error) {
		mu.Lock()
		counts[file.Path]++
		mu.Unlock()
		entry := indexedFileAnalysis{path: file.Path, moduleKind: file.ModuleKind}
		switch filepath.Base(file.Path) {
		case "A.bas":
			entry.symbols = []intel.Symbol{{Name: "Main", Module: "A", Kind: "sub"}}
			entry.callSites = []calls.CallSite{{
				File: "src/modules/A.bas", Module: "A",
				Caller: &calls.Caller{Name: "Main", Kind: "sub", QualifiedName: "A.Main"},
				Callee: calls.Callee{Text: "Foo", BaseName: "Foo"},
				Range:  vbaast.Range{StartLine: 2, StartColumn: 4},
			}}
		case "B.bas":
			if strings.Contains(string(source), "Foo") {
				entry.symbols = []intel.Symbol{{
					Name: "Foo", Module: "B", Kind: "sub", Range: intel.Range{Start: intel.Position{Line: 2}},
				}}
			}
		}
		return entry, nil
	}
	index := newWorkspaceAnalysisIndex(root, config.Default(), parse, nil)

	got, err := index.queryResolvedCalls(workspaceCallQuery{CalleeBase: "foo"})
	if err != nil || len(got) != 1 || got[0].Resolution.Status != "unresolved" {
		t.Fatalf("initial calls = %+v, %v", got, err)
	}
	if err := os.WriteFile(b, []byte("Foo"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := index.updatePath(b); err != nil {
		t.Fatal(err)
	}
	got, err = index.queryResolvedCalls(workspaceCallQuery{Caller: "a.main", CalleeText: " FOO "})
	if err != nil || len(got) != 1 || got[0].Resolution.Status != "matched" {
		t.Fatalf("resolved calls = %+v, %v", got, err)
	}
	if candidate := got[0].Resolution.Candidates[0]; candidate.File != "src/modules/B.bas" || candidate.Line != 3 {
		t.Fatalf("candidate = %+v", candidate)
	}
	snapshot, err := index.callGraphSnapshot()
	if err != nil {
		t.Fatal(err)
	}
	impact, err := callgraph.AnalyzeSnapshot(snapshot, callgraph.Request{Target: "A.Main", Direction: callgraph.DirectionCallees, Depth: 1})
	if err != nil || len(impact.DirectCallees) != 1 || impact.DirectCallees[0].ID.QualifiedName != "B.Foo" {
		t.Fatalf("incremental impact snapshot = %+v, %v", impact, err)
	}
	if err := os.Remove(b); err != nil {
		t.Fatal(err)
	}
	if err := index.updatePath(b); err != nil {
		t.Fatal(err)
	}
	got, err = index.queryResolvedCalls(workspaceCallQuery{})
	if err != nil || len(got) != 1 || got[0].Resolution.Status != "unresolved" {
		t.Fatalf("calls after callee deletion = %+v, %v", got, err)
	}
	mu.Lock()
	defer mu.Unlock()
	if counts[a] != 1 || counts[b] != 2 {
		t.Fatalf("parse counts = %#v, caller must remain warm", counts)
	}
}

func TestWorkspaceAnalysisIndexCallPostingsExcludeModuleCallsFromCallerAndSort(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	for _, name := range []string{"B.bas", "A.bas"} {
		if err := os.WriteFile(filepath.Join(moduleDir, name), []byte(name), 0o644); err != nil {
			t.Fatal(err)
		}
	}
	parse := func(_ context.Context, file symbols.SourceFile, _ []byte) (indexedFileAnalysis, error) {
		module := strings.TrimSuffix(filepath.Base(file.Path), ".bas")
		caller := &calls.Caller{Name: "Run", Kind: "sub", QualifiedName: module + ".Run"}
		if module == "B" {
			caller = nil
		}
		return indexedFileAnalysis{
			path: file.Path,
			callSites: []calls.CallSite{{
				File: "src/modules/" + module + ".bas", Module: module, Caller: caller,
				Callee: calls.Callee{Text: "  Target ( 1 ) ", BaseName: "Target"},
				Range:  vbaast.Range{StartLine: 4, StartColumn: 2},
			}},
		}, nil
	}
	index := newWorkspaceAnalysisIndex(root, config.Default(), parse, nil)
	all, err := index.queryResolvedCalls(workspaceCallQuery{CalleeBase: "TARGET"})
	if err != nil || len(all) != 2 || all[0].File != "src/modules/A.bas" || all[1].File != "src/modules/B.bas" {
		t.Fatalf("sorted base-name calls = %+v, %v", all, err)
	}
	byText, err := index.queryResolvedCalls(workspaceCallQuery{CalleeText: "target ( 1 )"})
	if err != nil || len(byText) != 2 {
		t.Fatalf("normalized-text calls = %+v, %v", byText, err)
	}
	byCaller, err := index.queryResolvedCalls(workspaceCallQuery{Caller: "B.Run"})
	if err != nil || len(byCaller) != 0 {
		t.Fatalf("module-level call leaked into caller posting = %+v, %v", byCaller, err)
	}
}

func TestWorkspaceAnalysisIndexRejectsStaleDiskCandidateAfterOverlay(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(moduleDir, "Main.bas")
	if err := os.WriteFile(path, []byte("Initial"), 0o644); err != nil {
		t.Fatal(err)
	}
	entered := make(chan struct{})
	release := make(chan struct{})
	parse := func(_ context.Context, file symbols.SourceFile, source []byte) (indexedFileAnalysis, error) {
		name := string(source)
		if name == "Stale" {
			close(entered)
			<-release
		}
		return indexedFileAnalysis{
			path: file.Path, symbols: []intel.Symbol{{Name: name, Module: "Main", Kind: "sub"}},
			callSites: []calls.CallSite{{File: "src/modules/Main.bas", Callee: calls.Callee{Text: name, BaseName: name}}},
		}, nil
	}
	index := newWorkspaceAnalysisIndex(root, config.Default(), parse, nil)
	if err := index.waitReady(); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte("Stale"), 0o644); err != nil {
		t.Fatal(err)
	}
	done := make(chan error, 1)
	go func() { done <- index.updatePath(path) }()
	<-entered
	snapshot := intel.NewAnalysisSnapshot(intel.Document{Path: path, Source: "Fresh", ModuleKind: "standard"})
	defer snapshot.Retire()
	doc := snapshot.Document()
	index.setOverlay(doc, indexedFileAnalysis{
		symbols:   []intel.Symbol{{Name: "Fresh", Module: "Main", Kind: "sub"}},
		callSites: []calls.CallSite{{File: "src/modules/Main.bas", Callee: calls.Callee{Text: "Fresh", BaseName: "Fresh"}}},
	})
	close(release)
	if err := <-done; err != nil {
		t.Fatal(err)
	}
	if got, _ := index.searchExact("Fresh"); len(got) != 1 {
		t.Fatalf("overlay was overwritten: %+v", got)
	}
	callsResult, err := index.queryResolvedCalls(workspaceCallQuery{CalleeBase: "Fresh"})
	if err != nil || len(callsResult) != 1 {
		t.Fatalf("overlay calls were overwritten: %+v, %v", callsResult, err)
	}
}

func TestWorkspaceAnalysisIndexRejectsStaleDiskCandidateAfterDelete(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(moduleDir, "Main.bas")
	if err := os.WriteFile(path, []byte("Initial"), 0o644); err != nil {
		t.Fatal(err)
	}
	entered := make(chan struct{})
	release := make(chan struct{})
	parse := func(_ context.Context, file symbols.SourceFile, source []byte) (indexedFileAnalysis, error) {
		name := string(source)
		if name == "Stale" {
			close(entered)
			<-release
		}
		return indexedFileAnalysis{
			path: file.Path, symbols: []intel.Symbol{{Name: name, Module: "Main", Kind: "sub"}},
			callSites: []calls.CallSite{{File: "src/modules/Main.bas", Callee: calls.Callee{Text: name, BaseName: name}}},
		}, nil
	}
	index := newWorkspaceAnalysisIndex(root, config.Default(), parse, nil)
	if err := index.waitReady(); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte("Stale"), 0o644); err != nil {
		t.Fatal(err)
	}
	done := make(chan error, 1)
	go func() { done <- index.updatePath(path) }()
	<-entered
	if err := os.Remove(path); err != nil {
		t.Fatal(err)
	}
	if err := index.removePath(path); err != nil {
		t.Fatal(err)
	}
	close(release)
	if err := <-done; err != nil {
		t.Fatal(err)
	}
	if got, _ := index.searchContains(""); len(got) != 0 {
		t.Fatalf("stale symbols were published after deletion: %+v", got)
	}
	if got, err := index.queryResolvedCalls(workspaceCallQuery{}); err != nil || len(got) != 0 {
		t.Fatalf("stale calls were published after deletion: %+v, %v", got, err)
	}
}

func TestWorkspaceAnalysisIndexKeepsEffectiveEntryAfterParseFailure(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(moduleDir, "Main.bas")
	if err := os.WriteFile(path, []byte("Working"), 0o644); err != nil {
		t.Fatal(err)
	}
	parse := func(_ context.Context, file symbols.SourceFile, source []byte) (indexedFileAnalysis, error) {
		name := string(source)
		if name == "Broken" {
			return indexedFileAnalysis{}, os.ErrInvalid
		}
		return indexedFileAnalysis{
			path: file.Path, symbols: []intel.Symbol{{Name: name, Module: "Main", Kind: "sub"}},
			callSites: []calls.CallSite{{File: "src/modules/Main.bas", Callee: calls.Callee{Text: name, BaseName: name}}},
		}, nil
	}
	index := newWorkspaceAnalysisIndex(root, config.Default(), parse, nil)
	if err := index.waitReady(); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte("Broken"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := index.updatePath(path); err != os.ErrInvalid {
		t.Fatalf("parse error = %v, want %v", err, os.ErrInvalid)
	}
	if got, _ := index.searchExact("Working"); len(got) != 1 {
		t.Fatalf("effective symbols changed after parse failure: %+v", got)
	}
	if got, err := index.queryResolvedCalls(workspaceCallQuery{CalleeBase: "Working"}); err != nil || len(got) != 1 {
		t.Fatalf("effective calls changed after parse failure: %+v, %v", got, err)
	}
}

func BenchmarkWorkspaceAnalysisIndexWarmCallQuery(b *testing.B) {
	root := b.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		b.Fatal(err)
	}
	path := filepath.Join(moduleDir, "Main.bas")
	if err := os.WriteFile(path, []byte("benchmark"), 0o644); err != nil {
		b.Fatal(err)
	}
	parse := func(_ context.Context, file symbols.SourceFile, _ []byte) (indexedFileAnalysis, error) {
		return indexedFileAnalysis{
			path:    file.Path,
			symbols: []intel.Symbol{{Name: "Target", Module: "Main", Kind: "sub"}},
			callSites: []calls.CallSite{{
				File: "src/modules/Main.bas", Module: "Main",
				Caller: &calls.Caller{Name: "Run", Kind: "sub", QualifiedName: "Main.Run"},
				Callee: calls.Callee{Text: "Target", BaseName: "Target"},
				Range:  vbaast.Range{StartLine: 2},
			}},
		}, nil
	}
	index := newWorkspaceAnalysisIndex(root, config.Default(), parse, nil)
	if err := index.waitReady(); err != nil {
		b.Fatal(err)
	}

	b.ReportAllocs()
	b.ResetTimer()
	for range b.N {
		got, err := index.queryResolvedCalls(workspaceCallQuery{})
		if err != nil || len(got) != 1 {
			b.Fatalf("query = %+v, %v", got, err)
		}
	}
}
