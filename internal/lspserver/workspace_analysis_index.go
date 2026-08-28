package lspserver

import (
	"context"
	"crypto/sha256"
	"encoding/hex"
	"errors"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"sync"
	"time"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/callgraph"
	"github.com/harumiWeb/xlflow/internal/vba/calls"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
)

// workspaceAnalysisIndex keeps the on-disk workspace and open-document state as
// independent layers. Only the effective layer participates in postings.
//
// All mutation is path-scoped: replacing a file removes and re-adds only that
// file's references. The index never reparses unaffected source files.
type workspaceAnalysisIndex struct {
	root   string
	config config.Config
	parse  func(context.Context, symbols.SourceFile, []byte) (indexedFileAnalysis, error)
	// parseDeclarations is the lightweight phase used by the server startup
	// pipeline.  It is optional so the direct index tests and compatibility
	// callers can continue to provide one combined parser.
	parseDeclarations func(context.Context, symbols.SourceFile, []byte) (indexedFileAnalysis, error)
	log               func(fileCount int, started time.Time, err error)
	logDeclarations   func(fileCount int, started time.Time, err error)
	performance       *performanceRecorder
	// nonBlockingQueries allows LSP requests to observe the coherently indexed
	// subset while the initial background scan is still running. Direct index
	// users retain the historical wait-for-ready behavior by default.
	nonBlockingQueries bool

	mu          sync.RWMutex
	startOnce   sync.Once
	lifecycleMu sync.Mutex
	started     bool
	stopped     bool
	ready       chan struct{}
	readyOnce   sync.Once
	readyErr    error
	disk        map[string]indexedFileAnalysis
	overlays    map[string]indexedFileAnalysis
	// pending masks both the saved file and the last published overlay while a
	// newer in-memory document generation is being analyzed.  Workspace
	// queries therefore remain conservative instead of returning stale data.
	pending    map[string]uint64
	incomplete map[string]bool
	effective  map[string]indexedFileAnalysis
	generation map[string]uint64
	diskParses map[string]*diskParse
	// revision changes only when a complete effective entry is published or
	// removed. Pending overlays make projectSnapshot incomplete and therefore
	// cannot seed project-aware diagnostics.
	revision   uint64
	qualified  map[string][]symbolRef
	all        []symbolRef
	allCalls   []callRef
	byCaller   map[string][]callRef
	byBaseName map[string][]callRef
	byText     map[string][]callRef
	projectMu  sync.Mutex
	// projectDependencies retains only compact procedure/dependency metadata;
	// the previous full IR/CFG snapshot is intentionally not retained here.
	projectDependencies projectDependencyView
	// Snapshot preparation reuses the resolver index and each document's
	// overlay independently of the monotonically increasing workspace revision.
	snapshotResolverCache dependencyCache[procedureir.Resolver]
	snapshotViewCache     dependencyCache[procedureir.ResolvedDocumentView]
	declarations          *workspaceDeclarationIndex
	initialWorkers        int
	semanticWorkers       int
	initialCtx            context.Context
	initialCancel         context.CancelFunc
	initialWG             sync.WaitGroup
	stopOnce              sync.Once
}

type diskParse struct {
	generation uint64
	cancel     context.CancelFunc
}

type diskRefresh struct {
	key        string
	path       string
	generation uint64
	ctx        context.Context
	cancel     context.CancelFunc
	active     *diskParse
}

// initialSource is the immutable source generation shared by the declaration
// and semantic startup phases.  Keeping one source/version pair prevents the
// two layers from accidentally publishing results parsed from different reads
// when a watcher notification is still in flight.
type initialSource struct {
	source  []byte
	version string
}

type indexedFileAnalysis struct {
	path                  string
	version               string
	moduleKind            string
	declarationIncomplete bool
	source                string
	symbols               []intel.Symbol
	callSites             []calls.CallSite
	typeReferences        []calls.TypeReference
	procedureIR           procedureir.DocumentIR
	controlFlow           vbacfg.Document
	procedureCatalog      intel.ProcedureCatalog
}

// indexedAnalysisIncomplete reports whether an indexed entry was produced
// from parser recovery.  Recovered IR is still useful for document-local
// features, but it cannot prove project-wide negative resolutions: a missing
// declaration may simply be hidden by the parser error.  Keep the entry in
// the effective index for best-effort queries while making the project
// snapshot fail open until a clean parse is published.
func indexedAnalysisIncomplete(entry indexedFileAnalysis) bool {
	ir := entry.procedureIR
	if ir.Parse.HasError || ir.Parse.HasMissing {
		return true
	}
	for _, declaration := range ir.Declarations {
		if declaration.Recovered || len(declaration.ConditionalBranches) != 0 {
			return true
		}
	}
	for _, procedure := range ir.Procedures {
		if procedure.Symbol.Recovered || len(procedure.Symbol.ConditionalBranches) != 0 {
			return true
		}
		for _, declaration := range procedure.Declarations {
			if declaration.Recovered || len(declaration.ConditionalBranches) != 0 {
				return true
			}
		}
		for _, statement := range procedure.Statements {
			if statement.Recovered {
				return true
			}
		}
		for _, expression := range procedure.Expressions {
			if expression.Recovered {
				return true
			}
		}
		for _, event := range procedure.RaiseEvents {
			if event.Recovered || len(event.ConditionalBranches) != 0 {
				return true
			}
		}
	}
	return false
}

type symbolRef struct {
	path  string
	index int
}

type callRef struct {
	path  string
	index int
}

// workspaceCallQuery selects raw call sites through one optional secondary
// posting. An empty query selects every effective workspace call.
type workspaceCallQuery struct {
	Caller     string
	CallerKind string
	CalleeBase string
	CalleeText string
}

func newWorkspaceAnalysisIndex(root string, cfg config.Config, parse func(context.Context, symbols.SourceFile, []byte) (indexedFileAnalysis, error), logInitial func(int, time.Time, error)) *workspaceAnalysisIndex {
	initialCtx, initialCancel := context.WithCancel(context.Background())
	return &workspaceAnalysisIndex{
		root: root, config: cfg, parse: parse, log: logInitial, ready: make(chan struct{}),
		disk: map[string]indexedFileAnalysis{}, overlays: map[string]indexedFileAnalysis{}, pending: map[string]uint64{}, incomplete: map[string]bool{}, effective: map[string]indexedFileAnalysis{},
		generation: map[string]uint64{}, diskParses: map[string]*diskParse{}, qualified: map[string][]symbolRef{},
		byCaller: map[string][]callRef{}, byBaseName: map[string][]callRef{}, byText: map[string][]callRef{},
		declarations: newWorkspaceDeclarationIndex(root, cfg), initialWorkers: 1, semanticWorkers: 1,
		initialCtx: initialCtx, initialCancel: initialCancel,
	}
}

func (x *workspaceAnalysisIndex) start() {
	x.lifecycleMu.Lock()
	defer x.lifecycleMu.Unlock()
	if x.stopped {
		return
	}
	x.startOnce.Do(func() {
		x.started = true
		x.initialWG.Add(1)
		go func() {
			defer x.initialWG.Done()
			x.buildInitial()
		}()
	})
}

func (x *workspaceAnalysisIndex) waitReady() error {
	x.start()
	<-x.ready
	x.mu.RLock()
	err := x.readyErr
	x.mu.RUnlock()
	return err
}

func (x *workspaceAnalysisIndex) queryReady() error {
	if !x.nonBlockingQueries {
		return x.waitReady()
	}
	x.start()
	select {
	case <-x.ready:
		x.mu.RLock()
		err := x.readyErr
		x.mu.RUnlock()
		return err
	default:
		return nil
	}
}

// waitDeclarationsReady is the lightweight readiness boundary used by
// workspace symbol queries.  waitReady remains the semantic compatibility
// boundary for project/call-graph consumers.
func (x *workspaceAnalysisIndex) waitDeclarationsReady() error {
	x.start()
	return x.declarations.waitReady()
}

func (x *workspaceAnalysisIndex) stop() {
	x.stopOnce.Do(func() {
		x.lifecycleMu.Lock()
		x.stopped = true
		started := x.started
		x.lifecycleMu.Unlock()
		if x.initialCancel != nil {
			x.initialCancel()
		}
		x.mu.Lock()
		for _, active := range x.diskParses {
			active.cancel()
		}
		x.mu.Unlock()
		x.initialWG.Wait()
		if !started {
			x.declarations.markReady(context.Canceled)
			x.mu.Lock()
			x.readyErr = context.Canceled
			x.mu.Unlock()
			x.readyOnce.Do(func() { close(x.ready) })
		}
	})
}

func (x *workspaceAnalysisIndex) closeReady() {
	x.readyOnce.Do(func() { close(x.ready) })
}

// completeLocked reports whether the effective index is safe for
// project-wide negative resolution. Callers must hold x.mu (for reading or
// writing) while invoking this helper. In non-blocking query mode the initial
// scan may still be publishing entries, so pending/incomplete state alone is
// insufficient; readiness and its discovery error are part of the same
// coherence check.
func (x *workspaceAnalysisIndex) completeLocked() bool {
	if len(x.pending) != 0 || len(x.incomplete) != 0 || !x.declarations.complete() {
		return false
	}
	select {
	case <-x.ready:
		return x.readyErr == nil
	default:
		return false
	}
}

func (x *workspaceAnalysisIndex) buildInitial() {
	started := time.Now()
	discovery := x.performance.start("workspace/index", performanceStageWorkspaceDiscovery, "background", x.root)
	files, err := symbols.DiscoverSourceFilesContext(x.initialCtx, symbols.Options{RootDir: x.root, Config: x.config})
	discovery.finish(len(files), 0, err)
	x.performance.addCounter(performanceCounterWorkspaceFilesDiscovered, uint64(len(files)), "workspace/index", performanceStageWorkspaceDiscovery, "background", x.root)
	if err == nil && x.parseDeclarations != nil {
		// The declaration phase is intentionally complete before semantic
		// preparation starts.  This makes the symbol index useful as soon as
		// possible and keeps heavyweight IR/CFG work from delaying it.
		declarationStarted := time.Now()
		sources := x.buildInitialDeclarations(files)
		err = x.initialCtx.Err()
		x.declarations.markReady(err)
		if x.logDeclarations != nil {
			x.logDeclarations(len(files), declarationStarted, err)
		}
		if err == nil {
			x.buildInitialSemantics(files, sources)
		}
	} else if err == nil {
		for _, file := range files {
			// Compatibility path for direct index users that provide one
			// combined parser.  The combined entry is projected into both
			// layers after it completes.
			_ = x.upsertDisk(file, true)
		}
		err = x.initialCtx.Err()
		x.declarations.markReady(err)
	} else {
		x.declarations.markReady(err)
	}
	if err == nil {
		err = x.initialCtx.Err()
	}
	x.mu.Lock()
	x.readyErr = err
	x.mu.Unlock()
	if x.log != nil {
		x.log(len(files), started, err)
	}
	x.closeReady()
}

func (x *workspaceAnalysisIndex) buildInitialDeclarations(files []symbols.SourceFile) map[string]initialSource {
	workers := max(1, x.initialWorkers)
	jobs := make(chan symbols.SourceFile)
	sources := make(map[string]initialSource, len(files))
	var sourcesMu sync.Mutex
	var wg sync.WaitGroup
	for i := 0; i < workers; i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			for file := range jobs {
				if x.initialCtx != nil {
					if err := x.initialCtx.Err(); err != nil {
						return
					}
				}
				source, readErr := os.ReadFile(file.Path)
				if readErr != nil {
					x.declarations.recordInitialFailure(file.Path)
					continue
				}
				key := symbolFileKey(file.Path)
				version := sourceVersion(source)
				if key != "" {
					sourcesMu.Lock()
					sources[key] = initialSource{source: append([]byte(nil), source...), version: version}
					sourcesMu.Unlock()
				}
				entry, parseErr := x.parseDeclarations(x.initialCtx, file, source)
				if parseErr != nil {
					x.declarations.recordInitialFailure(file.Path)
					continue
				}
				entry.path = file.Path
				entry.version = version
				entry.moduleKind = file.ModuleKind
				x.declarations.setDisk(file.Path, entry, 0, true)
			}
		}()
	}
sendJobs:
	for _, file := range files {
		select {
		case jobs <- file:
		case <-x.initialCtx.Done():
			break sendJobs
		}
	}
	close(jobs)
	wg.Wait()
	x.declarations.sortAll()
	return sources
}

func (x *workspaceAnalysisIndex) buildInitialSemantics(files []symbols.SourceFile, sources map[string]initialSource) {
	workers := max(1, x.semanticWorkers)
	jobs := make(chan symbols.SourceFile)
	var wg sync.WaitGroup
	for i := 0; i < workers; i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			for file := range jobs {
				if err := x.initialCtx.Err(); err != nil {
					return
				}
				initial, ok := sources[symbolFileKey(file.Path)]
				if !ok {
					x.markInitialSemanticFailure(file.Path)
					continue
				}
				entry, parseErr := x.parse(x.initialCtx, file, initial.source)
				if parseErr != nil {
					x.markInitialSemanticFailure(file.Path)
					continue
				}
				entry.path = file.Path
				entry.version = initial.version
				entry.moduleKind = file.ModuleKind
				entry.source = string(initial.source)
				x.publishInitialSemantic(file.Path, entry)
			}
		}()
	}
sendJobs:
	for _, file := range files {
		select {
		case jobs <- file:
		case <-x.initialCtx.Done():
			break sendJobs
		}
	}
	close(jobs)
	wg.Wait()
}

func (x *workspaceAnalysisIndex) publishInitialSemantic(path string, entry indexedFileAnalysis) {
	key := symbolFileKey(path)
	if key == "" {
		return
	}
	x.mu.Lock()
	defer x.mu.Unlock()
	if x.generation[key] != 0 || x.pending[key] != 0 {
		return
	}
	x.disk[key] = entry
	if indexedAnalysisIncomplete(entry) {
		x.incomplete[key] = true
	} else {
		delete(x.incomplete, key)
	}
	if _, open := x.overlays[key]; !open {
		x.replaceEffectiveLocked(key, entry)
		x.revision++
	}
	if x.parseDeclarations == nil || len(entry.symbols) > 0 {
		// Once the full semantic parse is ready, keep the declaration query layer
		// in sync so closed-file lookups retain locals, constants, and labels.
		x.declarations.setEffectiveFromEntry(path, entry)
	}
}

func (x *workspaceAnalysisIndex) markInitialSemanticFailure(path string) {
	key := symbolFileKey(path)
	if key == "" {
		return
	}
	x.mu.Lock()
	defer x.mu.Unlock()
	if x.generation[key] != 0 || x.pending[key] != 0 {
		return
	}
	x.incomplete[key] = true
}

// updatePath accepts either a watcher create/change/delete notification. The
// current filesystem state wins over the notification kind, making duplicate
// and out-of-order notifications safe.
func (x *workspaceAnalysisIndex) updatePath(path string) error {
	file, included, err := symbols.SourceFileForPath(x.root, x.config, path)
	if err != nil {
		return err
	}
	if !included {
		return x.removePath(path)
	}
	return x.upsertDisk(file, false)
}

func (x *workspaceAnalysisIndex) upsertDisk(file symbols.SourceFile, initial bool) error {
	key := symbolFileKey(file.Path)
	if key == "" {
		return nil
	}
	x.mu.Lock()
	observed := x.generation[key]
	if !initial {
		observed++
		x.generation[key] = observed
	}
	// An open document is authoritative. Its pending or published overlay masks
	// disk state, and didClose will perform a fresh disk read before restoring
	// the path. Avoid analyzing the same large module once from disk and again
	// from the editor snapshot.
	if x.pending[key] != 0 {
		x.mu.Unlock()
		return nil
	}
	if _, open := x.overlays[key]; open {
		x.mu.Unlock()
		return nil
	}
	if active := x.diskParses[key]; active != nil {
		active.cancel()
	}
	ctx, cancel := context.WithCancel(context.Background())
	active := &diskParse{generation: observed, cancel: cancel}
	x.diskParses[key] = active
	// A refresh is not a coherent project view until its source has been
	// parsed and published.  Keep the previous effective entry available for
	// local queries, but suppress project-negative diagnostics while this parse
	// is in flight.
	x.incomplete[key] = true
	x.mu.Unlock()
	defer func() {
		cancel()
		x.mu.Lock()
		if x.diskParses[key] == active {
			delete(x.diskParses, key)
		}
		x.mu.Unlock()
	}()

	source, err := os.ReadFile(file.Path)
	if err != nil {
		if errors.Is(err, os.ErrNotExist) {
			return x.removePath(file.Path)
		}
		x.markDiskIncomplete(key, observed)
		if x.parseDeclarations != nil {
			declarationGeneration := x.declarations.beginDisk(file.Path)
			x.declarations.markIncomplete(file.Path, declarationGeneration)
		}
		return err
	}
	version := sourceVersion(source)
	x.mu.RLock()
	current, exists := x.disk[key]
	x.mu.RUnlock()
	currentIncomplete := indexedAnalysisIncomplete(current)
	if exists && current.version == version && current.moduleKind == file.ModuleKind &&
		(x.parseDeclarations == nil || x.declarations.matchesVersion(file.Path, version, file.ModuleKind)) {
		x.mu.Lock()
		if x.generation[key] == observed {
			if currentIncomplete {
				x.incomplete[key] = true
			} else {
				delete(x.incomplete, key)
			}
		}
		x.mu.Unlock()
		return nil
	}
	declarationGeneration := uint64(0)
	if x.parseDeclarations != nil {
		declarationGeneration = x.declarations.beginDisk(file.Path)
		declarationEntry, declarationErr := x.parseDeclarations(ctx, file, source)
		if declarationErr != nil {
			if errors.Is(declarationErr, context.Canceled) && ctx.Err() != nil {
				return nil
			}
			x.markDiskIncomplete(key, observed)
			x.declarations.markIncomplete(file.Path, declarationGeneration)
			return declarationErr
		}
		declarationEntry.path = file.Path
		declarationEntry.version = version
		declarationEntry.moduleKind = file.ModuleKind
		if !x.declarations.setDisk(file.Path, declarationEntry, declarationGeneration, false) {
			return nil
		}
	}
	entry, err := x.parse(ctx, file, source)
	if err != nil {
		if errors.Is(err, context.Canceled) && ctx.Err() != nil {
			return nil
		}
		x.markDiskIncomplete(key, observed)
		return err
	}
	entry.version = version
	entryIncomplete := indexedAnalysisIncomplete(entry)

	x.mu.Lock()
	defer x.mu.Unlock()
	if x.generation[key] != observed {
		return nil
	}
	x.disk[key] = entry
	if entryIncomplete {
		x.incomplete[key] = true
	} else {
		delete(x.incomplete, key)
	}
	if _, open := x.overlays[key]; !open && x.pending[key] == 0 {
		x.replaceEffectiveLocked(key, entry)
		x.revision++
	}
	if x.parseDeclarations == nil || len(entry.symbols) > 0 {
		// The compatibility parser returns declarations and semantic artifacts
		// together. With the split pipeline, project the complete semantic symbol
		// set after it has replaced the declaration-only entry. An empty semantic
		// result preserves the declaration result, which is useful for callers
		// whose semantic parser intentionally omits symbols.
		x.declarations.setEffectiveFromEntry(file.Path, entry)
	}
	return nil
}

// markDiskIncomplete records a read/parse failure for the path.  The previous
// effective entry remains available to document-local queries for compatibility
// (and to avoid flickering symbols), but project-wide negative diagnostics must
// fail open while this path is incomplete.  Generation checks prevent a
// superseded refresh from marking a newer overlay or disk publication.
func (x *workspaceAnalysisIndex) markDiskIncomplete(key string, generation uint64) {
	x.mu.Lock()
	defer x.mu.Unlock()
	if x.generation[key] != generation || x.pending[key] != 0 {
		return
	}
	if _, open := x.overlays[key]; open {
		return
	}
	x.incomplete[key] = true
}

func (x *workspaceAnalysisIndex) removePath(path string) error {
	key := symbolFileKey(path)
	if key == "" {
		return nil
	}
	x.mu.Lock()
	x.generation[key]++
	if active := x.diskParses[key]; active != nil {
		active.cancel()
		delete(x.diskParses, key)
	}
	delete(x.disk, key)
	delete(x.incomplete, key)
	if _, open := x.overlays[key]; !open && x.pending[key] == 0 {
		_, existed := x.effective[key]
		x.removeEffectiveLocked(key)
		if existed {
			x.revision++
		}
	}
	x.mu.Unlock()
	x.declarations.removePath(path)
	return nil
}

func (x *workspaceAnalysisIndex) setOverlay(doc intel.Document, analysis indexedFileAnalysis) {
	key := documentSymbolKey(doc)
	if key == "" {
		return
	}
	analysis.path = doc.Path
	analysis.version = documentVersion(doc)
	analysis.moduleKind = doc.ModuleKind
	analysisIncomplete := indexedAnalysisIncomplete(analysis)
	x.mu.Lock()
	x.generation[key]++
	if active := x.diskParses[key]; active != nil {
		active.cancel()
		delete(x.diskParses, key)
	}
	delete(x.pending, key)
	if analysisIncomplete {
		x.incomplete[key] = true
	} else {
		delete(x.incomplete, key)
	}
	x.overlays[key] = analysis
	x.replaceEffectiveLocked(key, analysis)
	x.revision++
	x.mu.Unlock()
	x.declarations.setOverlay(doc, analysis)
}

// beginOverlay reserves an open-document generation and removes any older
// effective entry. The returned analysis is the entry that workspace queries
// observed immediately before the reservation (either disk or overlay), and
// is used only to compare procedure signatures after publication.
func (x *workspaceAnalysisIndex) beginOverlay(doc intel.Document, generation uint64) (indexedFileAnalysis, bool) {
	key := documentSymbolKey(doc)
	if key == "" || generation == 0 {
		return indexedFileAnalysis{}, false
	}
	declarationPrevious, declarationExists := x.declarations.beginOverlay(doc, generation)
	x.mu.Lock()
	previous, ok := x.effective[key]
	x.generation[key]++
	if active := x.diskParses[key]; active != nil {
		active.cancel()
		delete(x.diskParses, key)
	}
	x.pending[key] = generation
	delete(x.incomplete, key)
	delete(x.overlays, key)
	x.removeEffectiveLocked(key)
	x.revision++
	x.mu.Unlock()
	if !ok {
		previous, ok = declarationPrevious, declarationExists
	}
	return previous, ok
}

// publishOverlay atomically publishes only the generation most recently
// reserved by beginOverlay. Superseded and close/reopen workers are rejected.
func (x *workspaceAnalysisIndex) publishOverlay(doc intel.Document, generation uint64, analysis indexedFileAnalysis) bool {
	declarationsPublished := x.publishOverlayDeclarations(doc, generation, analysis)
	semanticPublished := x.publishOverlaySemantic(doc, generation, analysis)
	return declarationsPublished && semanticPublished
}

func (x *workspaceAnalysisIndex) publishOverlayDeclarations(doc intel.Document, generation uint64, analysis indexedFileAnalysis) bool {
	return x.declarations.publishOverlay(doc, generation, analysis)
}

func (x *workspaceAnalysisIndex) publishOverlaySemantic(doc intel.Document, generation uint64, analysis indexedFileAnalysis) bool {
	key := documentSymbolKey(doc)
	if key == "" {
		return false
	}
	analysis.path = doc.Path
	analysis.version = documentVersion(doc)
	analysis.moduleKind = doc.ModuleKind
	analysisIncomplete := indexedAnalysisIncomplete(analysis)
	x.mu.Lock()
	defer x.mu.Unlock()
	if x.pending[key] != generation {
		return false
	}
	delete(x.pending, key)
	if analysisIncomplete {
		x.incomplete[key] = true
	} else {
		delete(x.incomplete, key)
	}
	x.overlays[key] = analysis
	x.replaceEffectiveLocked(key, analysis)
	x.revision++
	return true
}

// abandonOverlay ends a terminal overlay build without exposing saved symbols
// for the still-open buffer. An empty overlay marker keeps watcher disk updates
// masked until the next document generation or didClose.
func (x *workspaceAnalysisIndex) abandonOverlay(doc intel.Document, generation uint64) bool {
	x.declarations.abandonOverlay(doc, generation)
	key := documentSymbolKey(doc)
	if key == "" {
		return false
	}
	x.mu.Lock()
	defer x.mu.Unlock()
	if x.pending[key] != generation {
		return false
	}
	delete(x.pending, key)
	x.incomplete[key] = true
	x.overlays[key] = indexedFileAnalysis{
		path:       doc.Path,
		version:    documentVersion(doc),
		moduleKind: doc.ModuleKind,
	}
	x.removeEffectiveLocked(key)
	x.revision++
	return true
}

// abandonOverlaySemantic ends only the semantic part of an overlay build.
// Declaration publication is an independent readiness boundary and remains
// available for workspace symbol and navigation queries after semantic failure.
func (x *workspaceAnalysisIndex) abandonOverlaySemantic(doc intel.Document, generation uint64) bool {
	key := documentSymbolKey(doc)
	if key == "" {
		return false
	}
	x.mu.Lock()
	defer x.mu.Unlock()
	if x.pending[key] != generation {
		return false
	}
	delete(x.pending, key)
	x.incomplete[key] = true
	x.overlays[key] = indexedFileAnalysis{
		path:       doc.Path,
		version:    documentVersion(doc),
		moduleKind: doc.ModuleKind,
	}
	x.removeEffectiveLocked(key)
	x.revision++
	return true
}

// clearOverlay restores a freshly parsed disk entry. The path stays absent
// from the effective index for the entire refresh, so neither a stale cached
// disk entry nor the closing overlay can leak into concurrent queries. On a
// read or parse error the path remains absent.
func (x *workspaceAnalysisIndex) clearOverlay(path string) (indexedFileAnalysis, bool, error) {
	refresh := x.beginClearOverlay(path)
	if refresh == nil {
		return indexedFileAnalysis{}, false, nil
	}
	return x.finishClearOverlay(refresh)
}

func (x *workspaceAnalysisIndex) beginClearOverlay(path string) *diskRefresh {
	key := symbolFileKey(path)
	if key == "" {
		return nil
	}
	x.mu.Lock()
	x.generation[key]++
	observed := x.generation[key]
	if active := x.diskParses[key]; active != nil {
		active.cancel()
		delete(x.diskParses, key)
	}
	delete(x.pending, key)
	x.incomplete[key] = true
	delete(x.overlays, key)
	delete(x.disk, key)
	x.removeEffectiveLocked(key)
	x.revision++
	ctx, cancel := context.WithCancel(context.Background())
	active := &diskParse{generation: observed, cancel: cancel}
	x.diskParses[key] = active
	x.mu.Unlock()
	// Closing an open document must mask its declaration layer immediately as
	// well; otherwise a workspace symbol query can observe the stale overlay
	// while the fresh disk parse is still in flight.
	x.declarations.removePath(path)
	x.declarations.recordInitialFailure(path)
	return &diskRefresh{key: key, path: path, generation: observed, ctx: ctx, cancel: cancel, active: active}
}

func (x *workspaceAnalysisIndex) finishClearOverlay(refresh *diskRefresh) (indexedFileAnalysis, bool, error) {
	if refresh == nil {
		return indexedFileAnalysis{}, false, nil
	}
	defer func() {
		refresh.cancel()
		x.mu.Lock()
		if x.diskParses[refresh.key] == refresh.active {
			delete(x.diskParses, refresh.key)
		}
		x.mu.Unlock()
	}()

	file, included, err := symbols.SourceFileForPath(x.root, x.config, refresh.path)
	if err != nil {
		return indexedFileAnalysis{}, false, err
	}
	if !included {
		x.completeClearWithoutEntry(refresh)
		x.declarations.removePath(refresh.path)
		return indexedFileAnalysis{}, false, nil
	}
	source, err := os.ReadFile(file.Path)
	if err != nil {
		if errors.Is(err, os.ErrNotExist) {
			x.completeClearWithoutEntry(refresh)
			x.declarations.removePath(refresh.path)
			return indexedFileAnalysis{}, false, nil
		}
		return indexedFileAnalysis{}, false, err
	}
	entry, err := x.parse(refresh.ctx, file, source)
	if err != nil {
		if errors.Is(err, context.Canceled) && refresh.ctx.Err() != nil {
			return indexedFileAnalysis{}, false, nil
		}
		return indexedFileAnalysis{}, false, err
	}
	entry.path = file.Path
	entry.version = sourceVersion(source)
	entry.moduleKind = file.ModuleKind
	entryIncomplete := indexedAnalysisIncomplete(entry)

	x.mu.Lock()
	defer x.mu.Unlock()
	if x.generation[refresh.key] != refresh.generation || x.diskParses[refresh.key] != refresh.active {
		return indexedFileAnalysis{}, false, nil
	}
	x.disk[refresh.key] = entry
	if entryIncomplete {
		x.incomplete[refresh.key] = true
	} else {
		delete(x.incomplete, refresh.key)
	}
	x.replaceEffectiveLocked(refresh.key, entry)
	x.revision++
	x.declarations.setEffectiveFromEntry(file.Path, entry)
	return entry, true, nil
}

func (x *workspaceAnalysisIndex) completeClearWithoutEntry(refresh *diskRefresh) {
	x.mu.Lock()
	defer x.mu.Unlock()
	if x.generation[refresh.key] != refresh.generation || x.diskParses[refresh.key] != refresh.active {
		return
	}
	delete(x.incomplete, refresh.key)
	x.revision++
}

// projectSnapshot returns a defensive, coherent project view. A snapshot is
// incomplete until the initial scan finishes and whenever a newer editor
// overlay is pending, a disk parse failed, or recovered IR was published.
// Callers must not publish project diagnostics from an incomplete view.
func (x *workspaceAnalysisIndex) projectSnapshot() intel.ProjectAnalysisSnapshot {
	return x.projectSnapshotClass("background")
}

func (x *workspaceAnalysisIndex) projectSnapshotClass(class string) intel.ProjectAnalysisSnapshot {
	x.start()
	snapshotMeasurement := x.performance.start("workspace/project", performanceStageProjectSnapshot, class, x.root)
	x.performance.addCounter(performanceCounterProjectSnapshotBuilds, 1, "workspace/project", performanceStageProjectSnapshot, class, x.root)
	x.mu.RLock()
	complete := x.completeLocked()
	revision := x.revision
	keys := make([]string, 0, len(x.effective))
	for key := range x.effective {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	rawEntries := make([]indexedFileAnalysis, 0, len(keys))
	symbolCapacity := len(x.all)
	for _, key := range keys {
		rawEntries = append(rawEntries, x.effective[key])
	}
	x.mu.RUnlock()

	entries := make([]indexedFileAnalysis, 0, len(rawEntries))
	resolverSymbols := make([]procedureir.ResolverSymbol, 0, symbolCapacity)
	graphSymbols := make([]callgraph.Symbol, 0, symbolCapacity)
	for _, raw := range rawEntries {
		entry := indexedFileAnalysis{
			path: raw.path, version: raw.version, source: raw.source, procedureIR: procedureir.Clone(raw.procedureIR),
			controlFlow: vbacfg.CloneDocument(raw.controlFlow),
			callSites:   cloneCallSites(raw.callSites), typeReferences: cloneTypeReferences(raw.typeReferences),
			symbols:          append([]intel.Symbol(nil), raw.symbols...),
			procedureCatalog: cloneProcedureCatalog(raw.procedureCatalog),
		}
		entries = append(entries, entry)
		file := workspaceDisplayPath(x.root, entry.path)
		for _, sym := range entry.symbols {
			resolverSymbols = append(resolverSymbols, procedureir.ResolverSymbol{
				Name: sym.Name, Type: sym.ReturnType, Module: sym.Module, ModuleKind: sym.ModuleKind, Kind: sym.Kind,
				Visibility: sym.Visibility, File: file, Line: sym.Range.Start.Line + 1, Parent: sym.Parent, IsArray: sym.IsArray,
				IsConst: procedureir.IsConstKind(sym.Kind),
			})
			graphSymbols = append(graphSymbols, callgraph.Symbol{
				Name: sym.Name, Kind: sym.Kind, Module: sym.Module, ModuleKind: sym.ModuleKind, File: file,
				Line: sym.Range.Start.Line + 1, Column: sym.Range.Start.Character + 1,
				EndLine: sym.Range.End.Line + 1, EndColumn: sym.Range.End.Character + 1,
				Parent: sym.Parent, Visibility: sym.Visibility, ReturnType: sym.ReturnType, Signature: sym.Detail,
			})
		}
	}

	resolverKey := resolverSymbolsFingerprint(resolverSymbols, complete)
	resolver, resolverErr, resolverHit := x.snapshotResolverCache.getOrBuildContext(context.Background(), resolverKey, func() (procedureir.Resolver, error) {
		resolverMeasurement := x.performance.start("workspace/project", performanceStageProjectResolver, class, x.root)
		x.performance.addCounter(performanceCounterResolutionResolverBuilds, 1, "workspace/project", performanceStageProjectResolver, class, x.root)
		value := procedureir.NewResolverWithCompleteness(resolverSymbols, complete)
		resolverMeasurement.finish(len(resolverSymbols), 0, nil)
		return value, nil
	})
	if resolverErr != nil {
		snapshotMeasurement.finish(0, 0, resolverErr)
		return intel.ProjectAnalysisSnapshot{Revision: revision, Complete: false}
	}
	if resolverHit {
		x.performance.addCounter(performanceCounterProjectCacheHits, 1, "workspace/project", performanceStageProjectResolver, class, x.root)
	} else {
		x.performance.addCounter(performanceCounterProjectCacheMisses, 1, "workspace/project", performanceStageProjectResolver, class, x.root)
		x.performance.addCounter(performanceCounterProjectCacheRebuilds, 1, "workspace/project", performanceStageProjectResolver, class, x.root)
	}
	viewMeasurement := x.performance.start("workspace/project", performanceStageProjectResolutionView, class, x.root)
	if !resolverHit {
		x.performance.addCounter(performanceCounterResolutionViewBuilds, 1, "workspace/project", performanceStageProjectResolutionView, class, x.root)
	}
	canonicalResolver, ok := resolver.(procedureir.SymbolResolver)
	if !ok {
		snapshotMeasurement.finish(0, 0, errors.New("cached workspace resolver has unexpected type"))
		return intel.ProjectAnalysisSnapshot{Revision: revision, Complete: false}
	}
	callResolver := calls.NewResolverFromProcedureIRResolver(canonicalResolver)
	viewMeasurement.finish(len(resolverSymbols), 0, nil)
	result := intel.ProjectAnalysisSnapshot{Revision: revision, Complete: complete}
	var sites []calls.CallSite
	var typeReferences []calls.TypeReference
	materializationMeasurement := x.performance.start("workspace/project", performanceStageResolutionMaterialize, class, x.root)
	for _, entry := range entries {
		viewKey := workspaceResolutionDocumentFingerprint(entry.path, entry.version, resolverKey, complete)
		resolution, viewErr, viewHit := x.snapshotViewCache.getOrBuildContext(context.Background(), viewKey, func() (procedureir.ResolvedDocumentView, error) {
			return procedureir.ResolveView(entry.procedureIR, resolver), nil
		})
		if viewErr != nil {
			materializationMeasurement.finish(0, 0, viewErr)
			snapshotMeasurement.finish(0, 0, viewErr)
			return intel.ProjectAnalysisSnapshot{Revision: revision, Complete: false}
		}
		if viewHit {
			x.performance.addCounter(performanceCounterProjectCacheHits, 1, "workspace/project", performanceStageProjectResolutionView, class, entry.path)
			x.performance.addCounter(performanceCounterProjectCacheReusedEntries, 1, "workspace/project", performanceStageProjectResolutionView, class, entry.path)
		} else {
			x.performance.addCounter(performanceCounterProjectCacheMisses, 1, "workspace/project", performanceStageProjectResolutionView, class, entry.path)
			x.performance.addCounter(performanceCounterProjectCacheRebuilds, 1, "workspace/project", performanceStageProjectResolutionView, class, entry.path)
			x.performance.addCounter(performanceCounterResolutionOverlayBuilds, 1, "workspace/project", performanceStageProjectResolutionView, class, entry.path)
		}
		result.Documents = append(result.Documents, intel.ProjectAnalysisDocument{
			IR: entry.procedureIR, Resolution: resolution, CFG: entry.controlFlow, Source: entry.source,
			Version: entry.version, ProcedureCatalog: entry.procedureCatalog,
		})
		sites = append(sites, entry.callSites...)
		typeReferences = append(typeReferences, entry.typeReferences...)
	}
	materializationMeasurement.finish(0, 0, nil)
	resolvedCalls := make([]calls.Call, len(sites))
	for i, site := range sites {
		resolvedCalls[i] = callResolver.Resolve(site)
	}
	sort.SliceStable(resolvedCalls, func(i, j int) bool { return callSiteLess(resolvedCalls[i].CallSite, resolvedCalls[j].CallSite) })
	sort.SliceStable(typeReferences, func(i, j int) bool {
		if typeReferences[i].File != typeReferences[j].File {
			return typeReferences[i].File < typeReferences[j].File
		}
		if typeReferences[i].Range.StartLine != typeReferences[j].Range.StartLine {
			return typeReferences[i].Range.StartLine < typeReferences[j].Range.StartLine
		}
		return typeReferences[i].Range.StartColumn < typeReferences[j].Range.StartColumn
	})
	result.CallGraph = callgraph.Snapshot{Symbols: graphSymbols, Calls: resolvedCalls, TypeReferences: typeReferences}
	snapshotMeasurement.finish(len(result.Documents), 0, nil)
	return result
}

func (x *workspaceAnalysisIndex) initialReady() bool {
	x.start()
	select {
	case <-x.ready:
		return true
	default:
		return false
	}
}

// projectChange advances the last complete snapshot and returns the affected
// transitive caller files. Incomplete snapshots never replace the baseline,
// so overlapping pending overlays are compared once the workspace becomes
// coherent again.
func (x *workspaceAnalysisIndex) projectChange() (intel.ProjectAnalysisSnapshot, []string) {
	return x.projectChangeClass("background")
}

func (x *workspaceAnalysisIndex) projectChangeClass(class string) (intel.ProjectAnalysisSnapshot, []string) {
	changeMeasurement := x.performance.start("workspace/project", performanceStageProjectChange, class, x.root)
	current := x.projectSnapshotClass(class)
	if !current.Complete {
		changeMeasurement.finish(0, 0, nil)
		return current, nil
	}
	x.projectMu.Lock()
	if x.projectDependencies.revision != 0 && current.Revision <= x.projectDependencies.revision {
		x.projectMu.Unlock()
		changeMeasurement.finish(0, 0, nil)
		return current, nil
	}
	if x.projectDependencies.revision == 0 {
		x.projectDependencies = buildProjectDependencyViewWithPerformanceClass(current, x.performance, class)
		x.projectDependencies.revision = current.Revision
		x.projectMu.Unlock()
		changeMeasurement.finish(0, 0, nil)
		return current, nil
	}
	dependencyMeasurement := x.performance.start("workspace/project", performanceStageDependencyUpdate, class, x.root)
	impacted := updateProjectDependencyView(&x.projectDependencies, current, x.performance, class)
	x.projectDependencies.revision = current.Revision
	x.projectMu.Unlock()
	dependencyMeasurement.finish(len(impacted), 0, nil)
	changeMeasurement.finish(len(impacted), 0, nil)
	return current, impacted
}

func cloneCallSites(in []calls.CallSite) []calls.CallSite {
	out := make([]calls.CallSite, len(in))
	for i := range in {
		out[i] = calls.CloneCallSite(in[i])
	}
	return out
}

func cloneTypeReferences(in []calls.TypeReference) []calls.TypeReference {
	out := append([]calls.TypeReference(nil), in...)
	for i := range out {
		if out[i].Caller != nil {
			caller := *out[i].Caller
			out[i].Caller = &caller
		}
	}
	return out
}

func (x *workspaceAnalysisIndex) searchContains(query string) ([]intel.Symbol, error) {
	x.start()
	return x.declarations.searchContains(query, x.nonBlockingQueries)
}

func (x *workspaceAnalysisIndex) symbolSnapshot() ([]intel.Symbol, error) {
	x.start()
	return x.declarations.symbolSnapshot(x.nonBlockingQueries)
}

func (x *workspaceAnalysisIndex) searchExact(name string) ([]intel.Symbol, error) {
	x.start()
	return x.declarations.searchExact(name, x.nonBlockingQueries)
}

func (x *workspaceAnalysisIndex) searchPrefix(prefix string) ([]intel.Symbol, error) {
	x.start()
	return x.declarations.searchPrefix(prefix, x.nonBlockingQueries)
}

func (x *workspaceAnalysisIndex) searchQualified(name string) ([]intel.Symbol, error) {
	x.start()
	return x.declarations.searchQualified(name, x.nonBlockingQueries)
}

func (x *workspaceAnalysisIndex) searchModule(name string) ([]intel.Symbol, error) {
	x.start()
	return x.declarations.searchModule(name, x.nonBlockingQueries)
}

func (x *workspaceAnalysisIndex) searchKind(kind string) ([]intel.Symbol, error) {
	x.start()
	return x.declarations.searchKind(kind, x.nonBlockingQueries)
}

// queryResolvedCalls snapshots both raw sites and the effective symbol set
// under one read lock, then performs resolution without holding the index.
// This guarantees each result observes one workspace revision while keeping
// potentially expensive resolution out of mutation critical sections.
func (x *workspaceAnalysisIndex) queryResolvedCalls(query workspaceCallQuery) ([]calls.Call, error) {
	if err := x.queryReady(); err != nil {
		return nil, err
	}

	caller := normalizeCallQuery(query.Caller)
	callerKind := normalizeCallQuery(query.CallerKind)
	baseName := normalizeCallQuery(query.CalleeBase)
	calleeText := normalizeCallText(query.CalleeText)

	x.mu.RLock()
	complete := x.completeLocked()
	refs := x.callRefsForQueryLocked(caller, baseName, calleeText)
	sites := make([]calls.CallSite, 0, len(refs))
	for _, ref := range refs {
		site, ok := x.callForRefLocked(ref)
		if !ok || !matchesCallQuery(site, caller, callerKind, baseName, calleeText) {
			continue
		}
		sites = append(sites, calls.CloneCallSite(site))
	}
	resolverSymbols := make([]procedureir.ResolverSymbol, 0, len(x.all))
	for _, ref := range x.all {
		sym, ok := x.symbolForRefLocked(ref)
		if !ok {
			continue
		}
		entry := x.effective[ref.path]
		resolverSymbols = append(resolverSymbols, procedureir.ResolverSymbol{
			Name:       sym.Name,
			Type:       sym.ReturnType,
			Module:     sym.Module,
			ModuleKind: sym.ModuleKind,
			Kind:       sym.Kind,
			Visibility: sym.Visibility,
			File:       workspaceDisplayPath(x.root, entry.path),
			Line:       sym.Range.Start.Line + 1, Parent: sym.Parent, IsArray: sym.IsArray,
			IsConst: procedureir.IsConstKind(sym.Kind),
		})
	}
	x.mu.RUnlock()

	resolver := calls.NewResolverFromProcedureIRResolver(procedureir.NewResolverWithCompleteness(resolverSymbols, complete))
	resolved := make([]calls.Call, len(sites))
	for i, site := range sites {
		resolved[i] = resolver.Resolve(site)
	}
	sort.SliceStable(resolved, func(i, j int) bool {
		return callSiteLess(resolved[i].CallSite, resolved[j].CallSite)
	})
	return resolved, nil
}

// callGraphSnapshot returns a coherent resolved view of the effective
// workspace. It is intentionally protocol-neutral so graph clients can reuse
// LSP overlays and incremental indexing without triggering another source
// parse. Call hierarchy currently uses queryResolvedCalls directly; impact and
// future graph features can consume this snapshot.
func (x *workspaceAnalysisIndex) callGraphSnapshot() (callgraph.Snapshot, error) {
	if err := x.queryReady(); err != nil {
		return callgraph.Snapshot{}, err
	}
	x.mu.RLock()
	complete := x.completeLocked()
	sites := make([]calls.CallSite, 0, len(x.allCalls))
	typeReferences := make([]calls.TypeReference, 0)
	for _, ref := range x.allCalls {
		site, ok := x.callForRefLocked(ref)
		if ok {
			sites = append(sites, calls.CloneCallSite(site))
		}
	}
	for _, entry := range x.effective {
		typeReferences = append(typeReferences, entry.typeReferences...)
	}
	resolverSymbols := make([]procedureir.ResolverSymbol, 0, len(x.all))
	graphSymbols := make([]callgraph.Symbol, 0, len(x.all))
	for _, ref := range x.all {
		sym, ok := x.symbolForRefLocked(ref)
		if !ok {
			continue
		}
		entry := x.effective[ref.path]
		file := workspaceDisplayPath(x.root, entry.path)
		resolverSymbols = append(resolverSymbols, procedureir.ResolverSymbol{Name: sym.Name, Type: sym.ReturnType, Module: sym.Module, ModuleKind: sym.ModuleKind, Kind: sym.Kind, Visibility: sym.Visibility, File: file, Line: sym.Range.Start.Line + 1, Parent: sym.Parent, IsArray: sym.IsArray, IsConst: procedureir.IsConstKind(sym.Kind)})
		graphSymbols = append(graphSymbols, callgraph.Symbol{
			Name: sym.Name, Kind: sym.Kind, Module: sym.Module, ModuleKind: sym.ModuleKind, File: file,
			Line: sym.Range.Start.Line + 1, Column: sym.Range.Start.Character + 1, EndLine: sym.Range.End.Line + 1, EndColumn: sym.Range.End.Character + 1,
			Parent: sym.Parent, Visibility: sym.Visibility, ReturnType: sym.ReturnType, Signature: sym.Detail,
		})
	}
	x.mu.RUnlock()
	resolver := calls.NewResolverFromProcedureIRResolver(procedureir.NewResolverWithCompleteness(resolverSymbols, complete))
	resolved := make([]calls.Call, len(sites))
	for i, site := range sites {
		resolved[i] = resolver.Resolve(site)
	}
	sort.SliceStable(resolved, func(i, j int) bool { return callSiteLess(resolved[i].CallSite, resolved[j].CallSite) })
	sort.SliceStable(typeReferences, func(i, j int) bool {
		if typeReferences[i].File != typeReferences[j].File {
			return typeReferences[i].File < typeReferences[j].File
		}
		if typeReferences[i].Range.StartLine != typeReferences[j].Range.StartLine {
			return typeReferences[i].Range.StartLine < typeReferences[j].Range.StartLine
		}
		return typeReferences[i].Range.StartColumn < typeReferences[j].Range.StartColumn
	})
	return callgraph.Snapshot{Symbols: graphSymbols, Calls: resolved, TypeReferences: typeReferences}, nil
}

func (x *workspaceAnalysisIndex) replaceEffectiveLocked(key string, entry indexedFileAnalysis) {
	x.removeEffectiveLocked(key)
	x.effective[key] = entry
	for i, sym := range entry.symbols {
		ref := symbolRef{path: key, index: i}
		x.addPostingLocked(x.qualified, nil, normalizeSymbolQuery(qualifiedSymbolName(sym)), ref)
		x.all = insertSortedRef(x.all, ref, x.refLessLocked)
	}
	for i, site := range entry.callSites {
		ref := callRef{path: key, index: i}
		x.allCalls = insertSortedCallRef(x.allCalls, ref, x.callRefLessLocked)
		if site.Caller != nil {
			x.addCallPostingLocked(x.byCaller, normalizeCallQuery(site.Caller.QualifiedName), ref)
		}
		x.addCallPostingLocked(x.byBaseName, normalizeCallQuery(site.Callee.BaseName), ref)
		x.addCallPostingLocked(x.byText, normalizeCallText(site.Callee.Text), ref)
	}
}

func (x *workspaceAnalysisIndex) removeEffectiveLocked(key string) {
	old, ok := x.effective[key]
	if !ok {
		return
	}
	for i, sym := range old.symbols {
		ref := symbolRef{path: key, index: i}
		x.removePostingLocked(x.qualified, nil, normalizeSymbolQuery(qualifiedSymbolName(sym)), ref)
		x.all = removeRef(x.all, ref)
	}
	for i, site := range old.callSites {
		ref := callRef{path: key, index: i}
		x.allCalls = removeCallRef(x.allCalls, ref)
		if site.Caller != nil {
			x.removeCallPostingLocked(x.byCaller, normalizeCallQuery(site.Caller.QualifiedName), ref)
		}
		x.removeCallPostingLocked(x.byBaseName, normalizeCallQuery(site.Callee.BaseName), ref)
		x.removeCallPostingLocked(x.byText, normalizeCallText(site.Callee.Text), ref)
	}
	delete(x.effective, key)
}

func (x *workspaceAnalysisIndex) addCallPostingLocked(postings map[string][]callRef, key string, ref callRef) {
	if key == "" {
		return
	}
	postings[key] = insertSortedCallRef(postings[key], ref, x.callRefLessLocked)
}

func (x *workspaceAnalysisIndex) removeCallPostingLocked(postings map[string][]callRef, key string, ref callRef) {
	if key == "" {
		return
	}
	refs := removeCallRef(postings[key], ref)
	if len(refs) == 0 {
		delete(postings, key)
		return
	}
	postings[key] = refs
}

func (x *workspaceAnalysisIndex) callRefsForQueryLocked(caller, baseName, calleeText string) []callRef {
	switch {
	case caller != "":
		return x.byCaller[caller]
	case baseName != "":
		return x.byBaseName[baseName]
	case calleeText != "":
		return x.byText[calleeText]
	default:
		return x.allCalls
	}
}

func (x *workspaceAnalysisIndex) callForRefLocked(ref callRef) (calls.CallSite, bool) {
	entry, ok := x.effective[ref.path]
	if !ok || ref.index < 0 || ref.index >= len(entry.callSites) {
		return calls.CallSite{}, false
	}
	return entry.callSites[ref.index], true
}

func (x *workspaceAnalysisIndex) callRefLessLocked(left, right callRef) bool {
	a, aok := x.callForRefLocked(left)
	b, bok := x.callForRefLocked(right)
	if !aok || !bok {
		if left.path != right.path {
			return left.path < right.path
		}
		return left.index < right.index
	}
	if callSiteLess(a, b) {
		return true
	}
	if callSiteLess(b, a) {
		return false
	}
	if left.path != right.path {
		return left.path < right.path
	}
	return left.index < right.index
}

func (x *workspaceAnalysisIndex) addPostingLocked(postings map[string][]symbolRef, keys *[]string, key string, ref symbolRef) {
	if key == "" {
		return
	}
	if _, exists := postings[key]; !exists && keys != nil {
		*keys = insertString(*keys, key)
	}
	postings[key] = insertSortedRef(postings[key], ref, x.refLessLocked)
}

func (x *workspaceAnalysisIndex) removePostingLocked(postings map[string][]symbolRef, keys *[]string, key string, ref symbolRef) {
	if key == "" {
		return
	}
	refs := removeRef(postings[key], ref)
	if len(refs) == 0 {
		delete(postings, key)
		if keys != nil {
			*keys = removeString(*keys, key)
		}
		return
	}
	postings[key] = refs
}

func (x *workspaceAnalysisIndex) symbolForRefLocked(ref symbolRef) (intel.Symbol, bool) {
	entry, ok := x.effective[ref.path]
	if !ok || ref.index < 0 || ref.index >= len(entry.symbols) {
		return intel.Symbol{}, false
	}
	return entry.symbols[ref.index], true
}

func (x *workspaceAnalysisIndex) refLessLocked(left, right symbolRef) bool {
	a, aok := x.symbolForRefLocked(left)
	b, bok := x.symbolForRefLocked(right)
	if !aok || !bok {
		return left.path < right.path
	}
	if a.File != b.File {
		return a.File < b.File
	}
	if a.Range.Start.Line != b.Range.Start.Line {
		return a.Range.Start.Line < b.Range.Start.Line
	}
	if a.Range.Start.Character != b.Range.Start.Character {
		return a.Range.Start.Character < b.Range.Start.Character
	}
	if a.Name != b.Name {
		return a.Name < b.Name
	}
	return left.path < right.path
}

func normalizeSymbolQuery(value string) string { return strings.ToLower(strings.TrimSpace(value)) }

func normalizeCallQuery(value string) string { return strings.ToLower(strings.TrimSpace(value)) }

func normalizeCallText(value string) string {
	return strings.ToLower(strings.Join(strings.Fields(value), " "))
}

func workspaceDisplayPath(root, path string) string {
	relative, err := filepath.Rel(root, path)
	if err == nil && relative != ".." && !strings.HasPrefix(relative, ".."+string(filepath.Separator)) {
		return filepath.ToSlash(relative)
	}
	return filepath.ToSlash(path)
}

func matchesCallQuery(site calls.CallSite, caller, callerKind, baseName, calleeText string) bool {
	if caller != "" && (site.Caller == nil || normalizeCallQuery(site.Caller.QualifiedName) != caller) {
		return false
	}
	if callerKind != "" && (site.Caller == nil || !matchingCallProcedureKinds(site.Caller.Kind, callerKind)) {
		return false
	}
	if baseName != "" && normalizeCallQuery(site.Callee.BaseName) != baseName {
		return false
	}
	return calleeText == "" || normalizeCallText(site.Callee.Text) == calleeText
}

func matchingCallProcedureKinds(actual, requested string) bool {
	actual = normalizeCallQuery(actual)
	requested = normalizeCallQuery(requested)
	if actual == requested {
		return true
	}
	return actual == "property" && (requested == "property_get" || requested == "property_let" || requested == "property_set")
}

func callSiteLess(a, b calls.CallSite) bool {
	if a.File != b.File {
		return a.File < b.File
	}
	if a.Range.StartLine != b.Range.StartLine {
		return a.Range.StartLine < b.Range.StartLine
	}
	if a.Range.StartColumn != b.Range.StartColumn {
		return a.Range.StartColumn < b.Range.StartColumn
	}
	if a.Callee.Text != b.Callee.Text {
		return a.Callee.Text < b.Callee.Text
	}
	aCaller, bCaller := "", ""
	if a.Caller != nil {
		aCaller = a.Caller.QualifiedName
	}
	if b.Caller != nil {
		bCaller = b.Caller.QualifiedName
	}
	return aCaller < bCaller
}

func qualifiedSymbolName(sym intel.Symbol) string {
	if strings.TrimSpace(sym.Module) == "" {
		return sym.Name
	}
	return sym.Module + "." + sym.Name
}

func documentVersion(doc intel.Document) string {
	if doc.Snapshot != nil && doc.Snapshot.Matches(doc) {
		return doc.Snapshot.SourceHash()
	}
	return sourceVersion([]byte(doc.Source))
}

func sourceVersion(source []byte) string {
	sum := sha256.Sum256(source)
	return hex.EncodeToString(sum[:])
}

func insertSortedRef(refs []symbolRef, ref symbolRef, less func(symbolRef, symbolRef) bool) []symbolRef {
	i := sort.Search(len(refs), func(i int) bool { return !less(refs[i], ref) })
	refs = append(refs, symbolRef{})
	copy(refs[i+1:], refs[i:])
	refs[i] = ref
	return refs
}

func removeRef(refs []symbolRef, ref symbolRef) []symbolRef {
	for i, candidate := range refs {
		if candidate == ref {
			copy(refs[i:], refs[i+1:])
			return refs[:len(refs)-1]
		}
	}
	return refs
}

func insertSortedCallRef(refs []callRef, ref callRef, less func(callRef, callRef) bool) []callRef {
	i := sort.Search(len(refs), func(i int) bool { return !less(refs[i], ref) })
	refs = append(refs, callRef{})
	copy(refs[i+1:], refs[i:])
	refs[i] = ref
	return refs
}

func removeCallRef(refs []callRef, ref callRef) []callRef {
	for i, candidate := range refs {
		if candidate == ref {
			copy(refs[i:], refs[i+1:])
			return refs[:len(refs)-1]
		}
	}
	return refs
}

func insertString(values []string, value string) []string {
	i := sort.SearchStrings(values, value)
	values = append(values, "")
	copy(values[i+1:], values[i:])
	values[i] = value
	return values
}

func removeString(values []string, value string) []string {
	i := sort.SearchStrings(values, value)
	if i < len(values) && values[i] == value {
		return append(values[:i], values[i+1:]...)
	}
	return values
}
