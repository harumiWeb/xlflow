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
	log    func(fileCount int, started time.Time, err error)
	// nonBlockingQueries allows LSP requests to observe the coherently indexed
	// subset while the initial background scan is still running. Direct index
	// users retain the historical wait-for-ready behavior by default.
	nonBlockingQueries bool

	mu        sync.RWMutex
	startOnce sync.Once
	ready     chan struct{}
	readyErr  error
	disk      map[string]indexedFileAnalysis
	overlays  map[string]indexedFileAnalysis
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
	revision            uint64
	exactName           map[string][]symbolRef
	qualified           map[string][]symbolRef
	moduleName          map[string][]symbolRef
	symbolKind          map[string][]symbolRef
	exactKeys           []string
	qualKeys            []string
	all                 []symbolRef
	allCalls            []callRef
	byCaller            map[string][]callRef
	byBaseName          map[string][]callRef
	byText              map[string][]callRef
	projectMu           sync.Mutex
	lastProjectSnapshot intel.ProjectAnalysisSnapshot
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

type indexedFileAnalysis struct {
	path           string
	version        string
	moduleKind     string
	source         string
	symbols        []intel.Symbol
	callSites      []calls.CallSite
	typeReferences []calls.TypeReference
	procedureIR    procedureir.DocumentIR
	controlFlow    vbacfg.Document
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
	return &workspaceAnalysisIndex{
		root: root, config: cfg, parse: parse, log: logInitial, ready: make(chan struct{}),
		disk: map[string]indexedFileAnalysis{}, overlays: map[string]indexedFileAnalysis{}, pending: map[string]uint64{}, incomplete: map[string]bool{}, effective: map[string]indexedFileAnalysis{},
		generation: map[string]uint64{}, diskParses: map[string]*diskParse{}, exactName: map[string][]symbolRef{}, qualified: map[string][]symbolRef{}, moduleName: map[string][]symbolRef{}, symbolKind: map[string][]symbolRef{},
		byCaller: map[string][]callRef{}, byBaseName: map[string][]callRef{}, byText: map[string][]callRef{},
	}
}

func (x *workspaceAnalysisIndex) start() {
	x.startOnce.Do(func() { go x.buildInitial() })
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

func (x *workspaceAnalysisIndex) buildInitial() {
	started := time.Now()
	files, err := symbols.DiscoverSourceFiles(symbols.Options{RootDir: x.root, Config: x.config})
	if err == nil {
		for _, file := range files {
			if err = x.upsertDisk(file, true); err != nil {
				break
			}
		}
	}
	x.mu.Lock()
	x.readyErr = err
	x.mu.Unlock()
	if x.log != nil {
		x.log(len(files), started, err)
	}
	close(x.ready)
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
		return err
	}
	version := sourceVersion(source)
	x.mu.RLock()
	current, exists := x.disk[key]
	x.mu.RUnlock()
	if exists && current.version == version && current.moduleKind == file.ModuleKind {
		return nil
	}
	entry, err := x.parse(ctx, file, source)
	if err != nil {
		if errors.Is(err, context.Canceled) && ctx.Err() != nil {
			return nil
		}
		return err
	}
	entry.version = version

	x.mu.Lock()
	defer x.mu.Unlock()
	if x.generation[key] != observed {
		return nil
	}
	x.disk[key] = entry
	if _, open := x.overlays[key]; !open && x.pending[key] == 0 {
		x.replaceEffectiveLocked(key, entry)
		x.revision++
	}
	return nil
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
	if _, open := x.overlays[key]; !open && x.pending[key] == 0 {
		_, existed := x.effective[key]
		x.removeEffectiveLocked(key)
		if existed {
			x.revision++
		}
	}
	x.mu.Unlock()
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
	x.mu.Lock()
	x.generation[key]++
	if active := x.diskParses[key]; active != nil {
		active.cancel()
		delete(x.diskParses, key)
	}
	delete(x.pending, key)
	delete(x.incomplete, key)
	x.overlays[key] = analysis
	x.replaceEffectiveLocked(key, analysis)
	x.revision++
	x.mu.Unlock()
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
	return previous, ok
}

// publishOverlay atomically publishes only the generation most recently
// reserved by beginOverlay. Superseded and close/reopen workers are rejected.
func (x *workspaceAnalysisIndex) publishOverlay(doc intel.Document, generation uint64, analysis indexedFileAnalysis) bool {
	key := documentSymbolKey(doc)
	if key == "" {
		return false
	}
	analysis.path = doc.Path
	analysis.version = documentVersion(doc)
	analysis.moduleKind = doc.ModuleKind
	x.mu.Lock()
	defer x.mu.Unlock()
	if x.pending[key] != generation {
		return false
	}
	delete(x.pending, key)
	delete(x.incomplete, key)
	x.overlays[key] = analysis
	x.replaceEffectiveLocked(key, analysis)
	x.revision++
	return true
}

// abandonOverlay ends a terminal overlay build without exposing saved symbols
// for the still-open buffer. An empty overlay marker keeps watcher disk updates
// masked until the next document generation or didClose.
func (x *workspaceAnalysisIndex) abandonOverlay(doc intel.Document, generation uint64) bool {
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
		return indexedFileAnalysis{}, false, nil
	}
	source, err := os.ReadFile(file.Path)
	if err != nil {
		if errors.Is(err, os.ErrNotExist) {
			x.completeClearWithoutEntry(refresh)
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

	x.mu.Lock()
	defer x.mu.Unlock()
	if x.generation[refresh.key] != refresh.generation || x.diskParses[refresh.key] != refresh.active {
		return indexedFileAnalysis{}, false, nil
	}
	x.disk[refresh.key] = entry
	delete(x.incomplete, refresh.key)
	x.replaceEffectiveLocked(refresh.key, entry)
	x.revision++
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
// overlay is pending. Callers must not publish project diagnostics from an
// incomplete view.
func (x *workspaceAnalysisIndex) projectSnapshot() intel.ProjectAnalysisSnapshot {
	x.start()
	x.mu.RLock()
	complete := len(x.pending) == 0 && len(x.incomplete) == 0
	select {
	case <-x.ready:
		complete = complete && x.readyErr == nil
	default:
		complete = false
	}
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
			path: raw.path, procedureIR: procedureir.Clone(raw.procedureIR),
			controlFlow: vbacfg.CloneDocument(raw.controlFlow),
			callSites:   cloneCallSites(raw.callSites), typeReferences: cloneTypeReferences(raw.typeReferences),
			symbols: append([]intel.Symbol(nil), raw.symbols...),
		}
		entries = append(entries, entry)
		file := workspaceDisplayPath(x.root, entry.path)
		for _, sym := range entry.symbols {
			resolverSymbols = append(resolverSymbols, procedureir.ResolverSymbol{
				Name: sym.Name, Module: sym.Module, ModuleKind: sym.ModuleKind, Kind: sym.Kind,
				Visibility: sym.Visibility, File: entry.procedureIR.Path, Line: sym.Range.Start.Line + 1,
			})
			graphSymbols = append(graphSymbols, callgraph.Symbol{
				Name: sym.Name, Kind: sym.Kind, Module: sym.Module, ModuleKind: sym.ModuleKind, File: file,
				Line: sym.Range.Start.Line + 1, Column: sym.Range.Start.Character + 1,
				EndLine: sym.Range.End.Line + 1, EndColumn: sym.Range.End.Character + 1,
				Parent: sym.Parent, Visibility: sym.Visibility, ReturnType: sym.ReturnType, Signature: sym.Detail,
			})
		}
	}

	resolver := procedureir.NewResolver(resolverSymbols)
	callResolverSymbols := make([]calls.ResolverSymbol, len(graphSymbols))
	for i, sym := range graphSymbols {
		callResolverSymbols[i] = calls.ResolverSymbol{
			Name: sym.Name, Module: sym.Module, ModuleKind: sym.ModuleKind, Kind: sym.Kind,
			Visibility: sym.Visibility, File: sym.File, Line: sym.Line,
		}
	}
	callResolver := calls.NewResolverFromSymbols(callResolverSymbols)
	result := intel.ProjectAnalysisSnapshot{Revision: revision, Complete: complete}
	var sites []calls.CallSite
	var typeReferences []calls.TypeReference
	for _, entry := range entries {
		resolvedIR := procedureir.Resolve(entry.procedureIR, resolver)
		result.Documents = append(result.Documents, intel.ProjectAnalysisDocument{IR: resolvedIR, CFG: entry.controlFlow})
		sites = append(sites, entry.callSites...)
		typeReferences = append(typeReferences, entry.typeReferences...)
	}
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
	current := x.projectSnapshot()
	if !current.Complete {
		return current, nil
	}
	x.projectMu.Lock()
	previous := x.lastProjectSnapshot
	if previous.Complete && current.Revision <= previous.Revision {
		x.projectMu.Unlock()
		return current, nil
	}
	x.lastProjectSnapshot = current
	x.projectMu.Unlock()
	if !previous.Complete {
		return current, nil
	}
	return current, projectImpactPaths(previous, current)
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
	if err := x.queryReady(); err != nil {
		return nil, err
	}
	query = normalizeSymbolQuery(query)
	x.mu.RLock()
	defer x.mu.RUnlock()
	refs := x.all
	if query == "" {
		return x.symbolsForRefsLocked(refs), nil
	}
	out := make([]intel.Symbol, 0)
	for _, ref := range refs {
		sym, ok := x.symbolForRefLocked(ref)
		if ok && (strings.Contains(strings.ToLower(sym.Name), query) || strings.Contains(strings.ToLower(qualifiedSymbolName(sym)), query)) {
			out = append(out, sym)
		}
	}
	return out, nil
}

func (x *workspaceAnalysisIndex) symbolSnapshot() ([]intel.Symbol, error) {
	if err := x.queryReady(); err != nil {
		return nil, err
	}
	x.mu.RLock()
	defer x.mu.RUnlock()
	return x.symbolsForRefsLocked(x.all), nil
}

func (x *workspaceAnalysisIndex) searchExact(name string) ([]intel.Symbol, error) {
	if err := x.queryReady(); err != nil {
		return nil, err
	}
	x.mu.RLock()
	defer x.mu.RUnlock()
	return x.symbolsForRefsLocked(x.exactName[normalizeSymbolQuery(name)]), nil
}

func (x *workspaceAnalysisIndex) searchPrefix(prefix string) ([]intel.Symbol, error) {
	if err := x.queryReady(); err != nil {
		return nil, err
	}
	prefix = normalizeSymbolQuery(prefix)
	x.mu.RLock()
	defer x.mu.RUnlock()
	if prefix == "" {
		return x.symbolsForRefsLocked(x.all), nil
	}
	start := sort.SearchStrings(x.exactKeys, prefix)
	buckets := make([][]symbolRef, 0)
	for i := start; i < len(x.exactKeys) && strings.HasPrefix(x.exactKeys[i], prefix); i++ {
		buckets = append(buckets, x.exactName[x.exactKeys[i]])
	}
	return x.mergePostingBucketsLocked(buckets), nil
}

func (x *workspaceAnalysisIndex) searchQualified(name string) ([]intel.Symbol, error) {
	if err := x.queryReady(); err != nil {
		return nil, err
	}
	x.mu.RLock()
	defer x.mu.RUnlock()
	return x.symbolsForRefsLocked(x.qualified[normalizeSymbolQuery(name)]), nil
}

func (x *workspaceAnalysisIndex) searchModule(name string) ([]intel.Symbol, error) {
	if err := x.queryReady(); err != nil {
		return nil, err
	}
	x.mu.RLock()
	defer x.mu.RUnlock()
	return x.symbolsForRefsLocked(x.moduleName[normalizeSymbolQuery(name)]), nil
}

func (x *workspaceAnalysisIndex) searchKind(kind string) ([]intel.Symbol, error) {
	if err := x.queryReady(); err != nil {
		return nil, err
	}
	x.mu.RLock()
	defer x.mu.RUnlock()
	return x.symbolsForRefsLocked(x.symbolKind[normalizeSymbolQuery(kind)]), nil
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
	refs := x.callRefsForQueryLocked(caller, baseName, calleeText)
	sites := make([]calls.CallSite, 0, len(refs))
	for _, ref := range refs {
		site, ok := x.callForRefLocked(ref)
		if !ok || !matchesCallQuery(site, caller, callerKind, baseName, calleeText) {
			continue
		}
		sites = append(sites, calls.CloneCallSite(site))
	}
	resolverSymbols := make([]calls.ResolverSymbol, 0, len(x.all))
	for _, ref := range x.all {
		sym, ok := x.symbolForRefLocked(ref)
		if !ok {
			continue
		}
		entry := x.effective[ref.path]
		resolverSymbols = append(resolverSymbols, calls.ResolverSymbol{
			Name:       sym.Name,
			Module:     sym.Module,
			ModuleKind: sym.ModuleKind,
			Kind:       sym.Kind,
			Visibility: sym.Visibility,
			File:       workspaceDisplayPath(x.root, entry.path),
			Line:       sym.Range.Start.Line + 1,
		})
	}
	x.mu.RUnlock()

	resolver := calls.NewResolverFromSymbols(resolverSymbols)
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
	resolverSymbols := make([]calls.ResolverSymbol, 0, len(x.all))
	graphSymbols := make([]callgraph.Symbol, 0, len(x.all))
	for _, ref := range x.all {
		sym, ok := x.symbolForRefLocked(ref)
		if !ok {
			continue
		}
		entry := x.effective[ref.path]
		file := workspaceDisplayPath(x.root, entry.path)
		resolverSymbols = append(resolverSymbols, calls.ResolverSymbol{Name: sym.Name, Module: sym.Module, ModuleKind: sym.ModuleKind, Kind: sym.Kind, Visibility: sym.Visibility, File: file, Line: sym.Range.Start.Line + 1})
		graphSymbols = append(graphSymbols, callgraph.Symbol{
			Name: sym.Name, Kind: sym.Kind, Module: sym.Module, ModuleKind: sym.ModuleKind, File: file,
			Line: sym.Range.Start.Line + 1, Column: sym.Range.Start.Character + 1, EndLine: sym.Range.End.Line + 1, EndColumn: sym.Range.End.Character + 1,
			Parent: sym.Parent, Visibility: sym.Visibility, ReturnType: sym.ReturnType, Signature: sym.Detail,
		})
	}
	x.mu.RUnlock()
	resolver := calls.NewResolverFromSymbols(resolverSymbols)
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
		x.addPostingLocked(x.exactName, &x.exactKeys, normalizeSymbolQuery(sym.Name), ref)
		x.addPostingLocked(x.qualified, &x.qualKeys, normalizeSymbolQuery(qualifiedSymbolName(sym)), ref)
		x.addPostingLocked(x.moduleName, nil, normalizeSymbolQuery(sym.Module), ref)
		x.addPostingLocked(x.symbolKind, nil, normalizeSymbolQuery(sym.Kind), ref)
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
		x.removePostingLocked(x.exactName, &x.exactKeys, normalizeSymbolQuery(sym.Name), ref)
		x.removePostingLocked(x.qualified, &x.qualKeys, normalizeSymbolQuery(qualifiedSymbolName(sym)), ref)
		x.removePostingLocked(x.moduleName, nil, normalizeSymbolQuery(sym.Module), ref)
		x.removePostingLocked(x.symbolKind, nil, normalizeSymbolQuery(sym.Kind), ref)
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

func (x *workspaceAnalysisIndex) symbolsForRefsLocked(refs []symbolRef) []intel.Symbol {
	out := make([]intel.Symbol, 0, len(refs))
	for _, ref := range refs {
		if sym, ok := x.symbolForRefLocked(ref); ok {
			out = append(out, sym)
		}
	}
	return out
}

func (x *workspaceAnalysisIndex) mergePostingBucketsLocked(buckets [][]symbolRef) []intel.Symbol {
	positions := make([]int, len(buckets))
	out := make([]intel.Symbol, 0)
	for {
		best := -1
		for i, bucket := range buckets {
			if positions[i] >= len(bucket) {
				continue
			}
			if best < 0 || x.refLessLocked(bucket[positions[i]], buckets[best][positions[best]]) {
				best = i
			}
		}
		if best < 0 {
			return out
		}
		ref := buckets[best][positions[best]]
		positions[best]++
		if symbol, ok := x.symbolForRefLocked(ref); ok {
			out = append(out, symbol)
		}
	}
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
