package lspserver

import (
	"sort"
	"strings"
	"sync"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
)

// workspaceDeclarationIndex owns the lightweight, query-oriented view of the
// workspace.  It deliberately stores no ProcedureIR, CFG, or call-site data;
// those artifacts remain owned by workspaceAnalysisIndex.
type workspaceDeclarationIndex struct {
	mu sync.RWMutex

	readyOnce sync.Once
	ready     chan struct{}
	readyErr  error

	disk       map[string]indexedFileAnalysis
	overlays   map[string]indexedFileAnalysis
	pending    map[string]uint64
	incomplete map[string]bool
	effective  map[string]indexedFileAnalysis
	generation map[string]uint64

	revision   uint64
	exactName  map[string][]symbolRef
	qualified  map[string][]symbolRef
	moduleName map[string][]symbolRef
	symbolKind map[string][]symbolRef
	exactKeys  []string
	qualKeys   []string
	all        []symbolRef

	root   string
	config config.Config
}

func newWorkspaceDeclarationIndex(root string, cfg config.Config) *workspaceDeclarationIndex {
	return &workspaceDeclarationIndex{
		root: root, config: cfg, ready: make(chan struct{}),
		disk: map[string]indexedFileAnalysis{}, overlays: map[string]indexedFileAnalysis{},
		pending: map[string]uint64{}, incomplete: map[string]bool{},
		effective: map[string]indexedFileAnalysis{}, generation: map[string]uint64{},
		exactName: map[string][]symbolRef{}, qualified: map[string][]symbolRef{},
		moduleName: map[string][]symbolRef{}, symbolKind: map[string][]symbolRef{},
	}
}

func (d *workspaceDeclarationIndex) markReady(err error) {
	d.readyOnce.Do(func() {
		d.mu.Lock()
		d.readyErr = err
		d.mu.Unlock()
		close(d.ready)
	})
}

func (d *workspaceDeclarationIndex) waitReady() error {
	<-d.ready
	d.mu.RLock()
	err := d.readyErr
	d.mu.RUnlock()
	return err
}

func (d *workspaceDeclarationIndex) queryReady(nonBlocking bool) error {
	if !nonBlocking {
		return d.waitReady()
	}
	select {
	case <-d.ready:
		d.mu.RLock()
		err := d.readyErr
		d.mu.RUnlock()
		return err
	default:
		return nil
	}
}

func (d *workspaceDeclarationIndex) completeLocked() bool {
	if len(d.pending) != 0 || len(d.incomplete) != 0 {
		return false
	}
	select {
	case <-d.ready:
		return d.readyErr == nil
	default:
		return false
	}
}

func (d *workspaceDeclarationIndex) complete() bool {
	d.mu.RLock()
	defer d.mu.RUnlock()
	return d.completeLocked()
}

func (d *workspaceDeclarationIndex) setDisk(filePath string, entry indexedFileAnalysis, generation uint64, initial bool) bool {
	key := symbolFileKey(filePath)
	if key == "" {
		return false
	}
	entry.path = filePath
	if entry.moduleKind == "" {
		entry.moduleKind = entryModuleKind(entry)
	}
	d.mu.Lock()
	defer d.mu.Unlock()
	if !initial && d.generation[key] != generation {
		return false
	}
	if initial && d.generation[key] != 0 && d.generation[key] != generation {
		return false
	}
	d.disk[key] = entry
	if entry.declarationIncomplete {
		d.incomplete[key] = true
	} else {
		delete(d.incomplete, key)
	}
	if _, open := d.overlays[key]; !open && d.pending[key] == 0 {
		d.replaceEffectiveLocked(key, entry)
		d.revision++
	}
	return true
}

func (d *workspaceDeclarationIndex) beginDisk(path string) uint64 {
	key := symbolFileKey(path)
	if key == "" {
		return 0
	}
	d.mu.Lock()
	defer d.mu.Unlock()
	d.generation[key]++
	generation := d.generation[key]
	if d.pending[key] == 0 {
		d.incomplete[key] = true
	}
	return generation
}

func (d *workspaceDeclarationIndex) matchesVersion(path, version, moduleKind string) bool {
	key := symbolFileKey(path)
	if key == "" {
		return false
	}
	d.mu.RLock()
	defer d.mu.RUnlock()
	entry, ok := d.disk[key]
	return ok && entry.version == version && entry.moduleKind == moduleKind && !d.incomplete[key]
}

// setEffectiveFromEntry is used by the legacy combined parser path, where a
// semantic entry also contains the declarations needed by this layer.
func (d *workspaceDeclarationIndex) setEffectiveFromEntry(filePath string, entry indexedFileAnalysis) {
	key := symbolFileKey(filePath)
	if key == "" {
		return
	}
	entry.path = filePath
	entry.declarationIncomplete = entry.declarationIncomplete || indexedAnalysisIncomplete(entry)
	d.mu.Lock()
	d.disk[key] = entry
	if entry.declarationIncomplete {
		d.incomplete[key] = true
	} else {
		delete(d.incomplete, key)
	}
	if _, open := d.overlays[key]; !open && d.pending[key] == 0 {
		d.replaceEffectiveLocked(key, entry)
		d.revision++
	}
	d.mu.Unlock()
}

func (d *workspaceDeclarationIndex) markIncomplete(filePath string, generation uint64) {
	key := symbolFileKey(filePath)
	if key == "" {
		return
	}
	d.mu.Lock()
	defer d.mu.Unlock()
	if d.generation[key] != generation || d.pending[key] != 0 {
		return
	}
	if _, open := d.overlays[key]; open {
		return
	}
	d.incomplete[key] = true
}

func (d *workspaceDeclarationIndex) recordInitialFailure(filePath string) {
	key := symbolFileKey(filePath)
	if key == "" {
		return
	}
	d.mu.Lock()
	defer d.mu.Unlock()
	if d.pending[key] != 0 {
		return
	}
	d.incomplete[key] = true
}

func (d *workspaceDeclarationIndex) removePath(path string) {
	key := symbolFileKey(path)
	if key == "" {
		return
	}
	d.mu.Lock()
	d.generation[key]++
	delete(d.disk, key)
	delete(d.incomplete, key)
	delete(d.pending, key)
	delete(d.overlays, key)
	if _, existed := d.effective[key]; existed {
		d.removeEffectiveLocked(key)
		d.revision++
	}
	d.mu.Unlock()
}

func (d *workspaceDeclarationIndex) setOverlay(doc intel.Document, entry indexedFileAnalysis) {
	key := documentSymbolKey(doc)
	if key == "" {
		return
	}
	entry.path, entry.version, entry.moduleKind = doc.Path, documentVersion(doc), doc.ModuleKind
	entry.declarationIncomplete = entry.declarationIncomplete || indexedAnalysisIncomplete(entry)
	d.mu.Lock()
	d.generation[key]++
	delete(d.pending, key)
	if entry.declarationIncomplete {
		d.incomplete[key] = true
	} else {
		delete(d.incomplete, key)
	}
	d.overlays[key] = entry
	d.replaceEffectiveLocked(key, entry)
	d.revision++
	d.mu.Unlock()
}

func (d *workspaceDeclarationIndex) beginOverlay(doc intel.Document, generation uint64) (indexedFileAnalysis, bool) {
	key := documentSymbolKey(doc)
	if key == "" || generation == 0 {
		return indexedFileAnalysis{}, false
	}
	d.mu.Lock()
	previous, ok := d.effective[key]
	d.generation[key]++
	d.pending[key] = generation
	delete(d.incomplete, key)
	delete(d.overlays, key)
	d.removeEffectiveLocked(key)
	d.revision++
	d.mu.Unlock()
	return previous, ok
}

func (d *workspaceDeclarationIndex) publishOverlay(doc intel.Document, generation uint64, entry indexedFileAnalysis) bool {
	key := documentSymbolKey(doc)
	if key == "" {
		return false
	}
	entry.path, entry.version, entry.moduleKind = doc.Path, documentVersion(doc), doc.ModuleKind
	entry.declarationIncomplete = entry.declarationIncomplete || indexedAnalysisIncomplete(entry)
	d.mu.Lock()
	defer d.mu.Unlock()
	if d.pending[key] != generation {
		return false
	}
	delete(d.pending, key)
	if entry.declarationIncomplete {
		d.incomplete[key] = true
	} else {
		delete(d.incomplete, key)
	}
	d.overlays[key] = entry
	d.replaceEffectiveLocked(key, entry)
	d.revision++
	return true
}

func (d *workspaceDeclarationIndex) abandonOverlay(doc intel.Document, generation uint64) bool {
	key := documentSymbolKey(doc)
	if key == "" {
		return false
	}
	d.mu.Lock()
	defer d.mu.Unlock()
	if d.pending[key] != generation {
		return false
	}
	delete(d.pending, key)
	d.incomplete[key] = true
	d.overlays[key] = indexedFileAnalysis{path: doc.Path, version: documentVersion(doc), moduleKind: doc.ModuleKind}
	d.removeEffectiveLocked(key)
	d.revision++
	return true
}

func (d *workspaceDeclarationIndex) searchContains(query string, nonBlocking bool) ([]intel.Symbol, error) {
	if err := d.queryReady(nonBlocking); err != nil {
		return nil, err
	}
	query = normalizeSymbolQuery(query)
	d.mu.RLock()
	defer d.mu.RUnlock()
	if query == "" {
		return d.symbolsForRefsLocked(d.all), nil
	}
	out := make([]intel.Symbol, 0)
	for _, ref := range d.all {
		sym, ok := d.symbolForRefLocked(ref)
		if ok && (strings.Contains(strings.ToLower(sym.Name), query) || strings.Contains(strings.ToLower(qualifiedSymbolName(sym)), query)) {
			out = append(out, sym)
		}
	}
	return out, nil
}

func (d *workspaceDeclarationIndex) symbolSnapshot(nonBlocking bool) ([]intel.Symbol, error) {
	if err := d.queryReady(nonBlocking); err != nil {
		return nil, err
	}
	d.mu.RLock()
	defer d.mu.RUnlock()
	return d.symbolsForRefsLocked(d.all), nil
}

func (d *workspaceDeclarationIndex) searchExact(name string, nonBlocking bool) ([]intel.Symbol, error) {
	if err := d.queryReady(nonBlocking); err != nil {
		return nil, err
	}
	d.mu.RLock()
	defer d.mu.RUnlock()
	return d.symbolsForRefsLocked(d.exactName[normalizeSymbolQuery(name)]), nil
}

func (d *workspaceDeclarationIndex) searchPrefix(prefix string, nonBlocking bool) ([]intel.Symbol, error) {
	if err := d.queryReady(nonBlocking); err != nil {
		return nil, err
	}
	prefix = normalizeSymbolQuery(prefix)
	d.mu.RLock()
	defer d.mu.RUnlock()
	if prefix == "" {
		return d.symbolsForRefsLocked(d.all), nil
	}
	start := sort.SearchStrings(d.exactKeys, prefix)
	buckets := make([][]symbolRef, 0)
	for i := start; i < len(d.exactKeys) && strings.HasPrefix(d.exactKeys[i], prefix); i++ {
		buckets = append(buckets, d.exactName[d.exactKeys[i]])
	}
	return d.mergePostingBucketsLocked(buckets), nil
}

func (d *workspaceDeclarationIndex) searchQualified(name string, nonBlocking bool) ([]intel.Symbol, error) {
	if err := d.queryReady(nonBlocking); err != nil {
		return nil, err
	}
	d.mu.RLock()
	defer d.mu.RUnlock()
	return d.symbolsForRefsLocked(d.qualified[normalizeSymbolQuery(name)]), nil
}

func (d *workspaceDeclarationIndex) searchModule(name string, nonBlocking bool) ([]intel.Symbol, error) {
	if err := d.queryReady(nonBlocking); err != nil {
		return nil, err
	}
	d.mu.RLock()
	defer d.mu.RUnlock()
	return d.symbolsForRefsLocked(d.moduleName[normalizeSymbolQuery(name)]), nil
}

func (d *workspaceDeclarationIndex) searchKind(kind string, nonBlocking bool) ([]intel.Symbol, error) {
	if err := d.queryReady(nonBlocking); err != nil {
		return nil, err
	}
	d.mu.RLock()
	defer d.mu.RUnlock()
	return d.symbolsForRefsLocked(d.symbolKind[normalizeSymbolQuery(kind)]), nil
}

func (d *workspaceDeclarationIndex) replaceEffectiveLocked(key string, entry indexedFileAnalysis) {
	d.removeEffectiveLocked(key)
	d.effective[key] = entry
	for i, sym := range entry.symbols {
		ref := symbolRef{path: key, index: i}
		d.addPostingLocked(d.exactName, &d.exactKeys, normalizeSymbolQuery(sym.Name), ref)
		d.addPostingLocked(d.qualified, &d.qualKeys, normalizeSymbolQuery(qualifiedSymbolName(sym)), ref)
		d.addPostingLocked(d.moduleName, nil, normalizeSymbolQuery(sym.Module), ref)
		d.addPostingLocked(d.symbolKind, nil, normalizeSymbolQuery(sym.Kind), ref)
		d.all = insertSortedRef(d.all, ref, d.refLessLocked)
	}
}

func (d *workspaceDeclarationIndex) removeEffectiveLocked(key string) {
	old, ok := d.effective[key]
	if !ok {
		return
	}
	for i, sym := range old.symbols {
		ref := symbolRef{path: key, index: i}
		d.removePostingLocked(d.exactName, &d.exactKeys, normalizeSymbolQuery(sym.Name), ref)
		d.removePostingLocked(d.qualified, &d.qualKeys, normalizeSymbolQuery(qualifiedSymbolName(sym)), ref)
		d.removePostingLocked(d.moduleName, nil, normalizeSymbolQuery(sym.Module), ref)
		d.removePostingLocked(d.symbolKind, nil, normalizeSymbolQuery(sym.Kind), ref)
		d.all = removeRef(d.all, ref)
	}
	delete(d.effective, key)
}

func (d *workspaceDeclarationIndex) addPostingLocked(postings map[string][]symbolRef, keys *[]string, key string, ref symbolRef) {
	if key == "" {
		return
	}
	if _, exists := postings[key]; !exists && keys != nil {
		*keys = insertString(*keys, key)
	}
	postings[key] = insertSortedRef(postings[key], ref, d.refLessLocked)
}

func (d *workspaceDeclarationIndex) removePostingLocked(postings map[string][]symbolRef, keys *[]string, key string, ref symbolRef) {
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

func (d *workspaceDeclarationIndex) symbolForRefLocked(ref symbolRef) (intel.Symbol, bool) {
	entry, ok := d.effective[ref.path]
	if !ok || ref.index < 0 || ref.index >= len(entry.symbols) {
		return intel.Symbol{}, false
	}
	return entry.symbols[ref.index], true
}

func (d *workspaceDeclarationIndex) symbolsForRefsLocked(refs []symbolRef) []intel.Symbol {
	out := make([]intel.Symbol, 0, len(refs))
	for _, ref := range refs {
		if sym, ok := d.symbolForRefLocked(ref); ok {
			out = append(out, sym)
		}
	}
	return out
}

func (d *workspaceDeclarationIndex) mergePostingBucketsLocked(buckets [][]symbolRef) []intel.Symbol {
	positions := make([]int, len(buckets))
	out := make([]intel.Symbol, 0)
	for {
		best := -1
		for i, bucket := range buckets {
			if positions[i] >= len(bucket) {
				continue
			}
			if best < 0 || d.refLessLocked(bucket[positions[i]], buckets[best][positions[best]]) {
				best = i
			}
		}
		if best < 0 {
			return out
		}
		ref := buckets[best][positions[best]]
		positions[best]++
		if symbol, ok := d.symbolForRefLocked(ref); ok {
			out = append(out, symbol)
		}
	}
}

func (d *workspaceDeclarationIndex) refLessLocked(left, right symbolRef) bool {
	a, aok := d.symbolForRefLocked(left)
	b, bok := d.symbolForRefLocked(right)
	if !aok || !bok {
		if left.path != right.path {
			return left.path < right.path
		}
		return left.index < right.index
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

func entryModuleKind(entry indexedFileAnalysis) string {
	if len(entry.symbols) != 0 {
		return entry.symbols[0].ModuleKind
	}
	return ""
}
