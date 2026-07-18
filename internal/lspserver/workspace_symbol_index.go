package lspserver

import (
	"crypto/sha256"
	"encoding/hex"
	"errors"
	"os"
	"sort"
	"strings"
	"sync"
	"time"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
)

// workspaceSymbolIndex keeps the on-disk workspace and open-document state as
// independent layers. Only the effective layer participates in postings.
//
// All mutation is path-scoped: replacing a file removes and re-adds only that
// file's references. The index never reparses unaffected source files.
type workspaceSymbolIndex struct {
	root   string
	config config.Config
	parse  func(symbols.SourceFile, []byte) (indexedFileSymbols, error)
	log    func(fileCount int, started time.Time, err error)

	mu         sync.RWMutex
	startOnce  sync.Once
	ready      chan struct{}
	readyErr   error
	disk       map[string]indexedFileSymbols
	overlays   map[string]indexedFileSymbols
	effective  map[string]indexedFileSymbols
	generation map[string]uint64
	exactName  map[string][]symbolRef
	qualified  map[string][]symbolRef
	moduleName map[string][]symbolRef
	symbolKind map[string][]symbolRef
	exactKeys  []string
	qualKeys   []string
	all        []symbolRef
}

type indexedFileSymbols struct {
	path       string
	version    string
	moduleKind string
	symbols    []intel.Symbol
}

type symbolRef struct {
	path  string
	index int
}

func newWorkspaceSymbolIndex(root string, cfg config.Config, parse func(symbols.SourceFile, []byte) (indexedFileSymbols, error), logInitial func(int, time.Time, error)) *workspaceSymbolIndex {
	return &workspaceSymbolIndex{
		root: root, config: cfg, parse: parse, log: logInitial, ready: make(chan struct{}),
		disk: map[string]indexedFileSymbols{}, overlays: map[string]indexedFileSymbols{}, effective: map[string]indexedFileSymbols{},
		generation: map[string]uint64{}, exactName: map[string][]symbolRef{}, qualified: map[string][]symbolRef{}, moduleName: map[string][]symbolRef{}, symbolKind: map[string][]symbolRef{},
	}
}

func (x *workspaceSymbolIndex) start() {
	x.startOnce.Do(func() { go x.buildInitial() })
}

func (x *workspaceSymbolIndex) waitReady() error {
	x.start()
	<-x.ready
	x.mu.RLock()
	err := x.readyErr
	x.mu.RUnlock()
	return err
}

func (x *workspaceSymbolIndex) buildInitial() {
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
func (x *workspaceSymbolIndex) updatePath(path string) error {
	file, included, err := symbols.SourceFileForPath(x.root, x.config, path)
	if err != nil {
		return err
	}
	if !included {
		return x.removePath(path)
	}
	return x.upsertDisk(file, false)
}

func (x *workspaceSymbolIndex) upsertDisk(file symbols.SourceFile, initial bool) error {
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
	x.mu.Unlock()

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
	entry, err := x.parse(file, source)
	if err != nil {
		return err
	}
	entry.version = version

	x.mu.Lock()
	defer x.mu.Unlock()
	if x.generation[key] != observed {
		return nil
	}
	x.disk[key] = entry
	if _, open := x.overlays[key]; !open {
		x.replaceEffectiveLocked(key, entry)
	}
	return nil
}

func (x *workspaceSymbolIndex) removePath(path string) error {
	key := symbolFileKey(path)
	if key == "" {
		return nil
	}
	x.mu.Lock()
	x.generation[key]++
	delete(x.disk, key)
	if _, open := x.overlays[key]; !open {
		x.removeEffectiveLocked(key)
	}
	x.mu.Unlock()
	return nil
}

func (x *workspaceSymbolIndex) setOverlay(doc intel.Document, symbols []intel.Symbol) {
	key := documentSymbolKey(doc)
	if key == "" {
		return
	}
	entry := indexedFileSymbols{path: doc.Path, version: documentVersion(doc), moduleKind: doc.ModuleKind, symbols: symbols}
	x.mu.Lock()
	x.generation[key]++
	x.overlays[key] = entry
	x.replaceEffectiveLocked(key, entry)
	x.mu.Unlock()
}

// clearOverlay restores a freshly read disk entry. Reloading instead of merely
// exposing the last watcher value closes editor/watcher ordering races.
func (x *workspaceSymbolIndex) clearOverlay(path string) error {
	key := symbolFileKey(path)
	if key == "" {
		return nil
	}
	x.mu.Lock()
	x.generation[key]++
	delete(x.overlays, key)
	if disk, ok := x.disk[key]; ok {
		x.replaceEffectiveLocked(key, disk)
	} else {
		x.removeEffectiveLocked(key)
	}
	x.mu.Unlock()
	return x.updatePath(path)
}

func (x *workspaceSymbolIndex) searchContains(query string) ([]intel.Symbol, error) {
	if err := x.waitReady(); err != nil {
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

func (x *workspaceSymbolIndex) searchExact(name string) ([]intel.Symbol, error) {
	if err := x.waitReady(); err != nil {
		return nil, err
	}
	x.mu.RLock()
	defer x.mu.RUnlock()
	return x.symbolsForRefsLocked(x.exactName[normalizeSymbolQuery(name)]), nil
}

func (x *workspaceSymbolIndex) searchPrefix(prefix string) ([]intel.Symbol, error) {
	if err := x.waitReady(); err != nil {
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

func (x *workspaceSymbolIndex) searchQualified(name string) ([]intel.Symbol, error) {
	if err := x.waitReady(); err != nil {
		return nil, err
	}
	x.mu.RLock()
	defer x.mu.RUnlock()
	return x.symbolsForRefsLocked(x.qualified[normalizeSymbolQuery(name)]), nil
}

func (x *workspaceSymbolIndex) searchModule(name string) ([]intel.Symbol, error) {
	if err := x.waitReady(); err != nil {
		return nil, err
	}
	x.mu.RLock()
	defer x.mu.RUnlock()
	return x.symbolsForRefsLocked(x.moduleName[normalizeSymbolQuery(name)]), nil
}

func (x *workspaceSymbolIndex) searchKind(kind string) ([]intel.Symbol, error) {
	if err := x.waitReady(); err != nil {
		return nil, err
	}
	x.mu.RLock()
	defer x.mu.RUnlock()
	return x.symbolsForRefsLocked(x.symbolKind[normalizeSymbolQuery(kind)]), nil
}

func (x *workspaceSymbolIndex) replaceEffectiveLocked(key string, entry indexedFileSymbols) {
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
}

func (x *workspaceSymbolIndex) removeEffectiveLocked(key string) {
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
	delete(x.effective, key)
}

func (x *workspaceSymbolIndex) addPostingLocked(postings map[string][]symbolRef, keys *[]string, key string, ref symbolRef) {
	if key == "" {
		return
	}
	if _, exists := postings[key]; !exists && keys != nil {
		*keys = insertString(*keys, key)
	}
	postings[key] = insertSortedRef(postings[key], ref, x.refLessLocked)
}

func (x *workspaceSymbolIndex) removePostingLocked(postings map[string][]symbolRef, keys *[]string, key string, ref symbolRef) {
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

func (x *workspaceSymbolIndex) symbolsForRefsLocked(refs []symbolRef) []intel.Symbol {
	out := make([]intel.Symbol, 0, len(refs))
	for _, ref := range refs {
		if sym, ok := x.symbolForRefLocked(ref); ok {
			out = append(out, sym)
		}
	}
	return out
}

func (x *workspaceSymbolIndex) mergePostingBucketsLocked(buckets [][]symbolRef) []intel.Symbol {
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

func (x *workspaceSymbolIndex) symbolForRefLocked(ref symbolRef) (intel.Symbol, bool) {
	entry, ok := x.effective[ref.path]
	if !ok || ref.index < 0 || ref.index >= len(entry.symbols) {
		return intel.Symbol{}, false
	}
	return entry.symbols[ref.index], true
}

func (x *workspaceSymbolIndex) refLessLocked(left, right symbolRef) bool {
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

func qualifiedSymbolName(sym intel.Symbol) string {
	if strings.TrimSpace(sym.Module) == "" {
		return sym.Name
	}
	return sym.Module + "." + sym.Name
}

func documentVersion(doc intel.Document) string {
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
