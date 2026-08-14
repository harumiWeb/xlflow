package intel

import (
	"context"
	"crypto/sha256"
	"fmt"
	"sort"
	"strings"

	staticrules "github.com/harumiWeb/xlflow/internal/staticanalysis/rules"
	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
	"github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/callgraph"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

type DiagnosticMode uint8

const (
	DiagnosticModeFull DiagnosticMode = iota
	DiagnosticModeFast
)

type ProcedureChangeSet struct {
	Ranges []Range
}

type DiagnosticRequest struct {
	Document      Document
	Mode          DiagnosticMode
	Changes       ProcedureChangeSet
	PreviousCache *DiagnosticCache
	Recorder      *analysisstats.Recorder
}

type DiagnosticResult struct {
	Diagnostics []Diagnostic
	Cache       *DiagnosticCache
}

// ProjectAnalysisDocument is an immutable, protocol-neutral document input for
// project-aware Full diagnostics. IR calls and source-derived constant values
// are resolved against the exact workspace snapshot represented by Revision;
// CFG is built from the same IR.
type ProjectAnalysisDocument struct {
	IR     procedureir.DocumentIR
	CFG    vbacfg.Document
	Source string
}

// ProjectAnalysisSnapshot is a coherent view of saved files and published
// editor overlays. Complete is false while any newer overlay is pending, so a
// project-aware diagnostic must not publish results from this snapshot.
type ProjectAnalysisSnapshot struct {
	Revision  uint64
	Complete  bool
	Documents []ProjectAnalysisDocument
	CallGraph callgraph.Snapshot
}

type ProcedureIdentity struct {
	CanonicalName string
	Kind          string
	Ordinal       int
}

type ProcedureCatalogEntry struct {
	Identity             ProcedureIdentity
	SignatureHash        [sha256.Size]byte
	SourceHash           [sha256.Size]byte
	Range                Range
	StartByte            int
	DeclarationStartByte int
	StartColumn          int
	EndByte              int
}

type ProcedureCatalog struct {
	Entries           []ProcedureCatalogEntry
	ModuleContextHash [sha256.Size]byte
	ConditionalHash   [sha256.Size]byte
	ReuseSafe         bool
}

type cachedProcedureDiagnostics struct {
	Entry       ProcedureCatalogEntry
	Diagnostics []Diagnostic
}

type DiagnosticCache struct {
	Catalog    ProcedureCatalog
	Procedures map[ProcedureIdentity]cachedProcedureDiagnostics
}

func (a Analyzer) DiagnosticsRequestContext(ctx context.Context, request DiagnosticRequest) DiagnosticResult {
	ctx = analysisstats.WithRecorder(ctx, request.Recorder)
	recorder := request.Recorder
	if recorder == nil {
		recorder = analysisstats.FromContext(ctx)
	}
	doc := request.Document
	var before ProcedureArtifactStats
	if doc.Snapshot != nil {
		before = doc.Snapshot.ProcedureArtifactStats()
	}
	defer func() {
		if recorder == nil || doc.Snapshot == nil {
			return
		}
		after := doc.Snapshot.ProcedureArtifactStats()
		recorder.Add("procedure_ir_builds", after.IRBuild-before.IRBuild)
		recorder.Add("procedure_ir_reuses", after.IRReuse-before.IRReuse)
		recorder.Add("cfg_builds", after.CFGBuild-before.CFGBuild)
		recorder.Add("cfg_reuses", after.CFGReuse-before.CFGReuse)
	}()
	a = a.withRequestWorkspaceResolution(ctx, []Document{doc})
	if ctx.Err() != nil {
		return DiagnosticResult{}
	}
	catalog := procedureCatalogForDocument(doc)
	if request.Mode != DiagnosticModeFast {
		diagnostics := a.diagnosticsFullContext(ctx, doc)
		if ctx.Err() != nil {
			return DiagnosticResult{}
		}
		return DiagnosticResult{Diagnostics: diagnostics, Cache: buildDiagnosticCache(catalog, diagnostics)}
	}
	// Property accessor contracts span sibling procedures. A procedure-sized
	// fast fragment cannot validate the group, so signature edits in any
	// Property accessor force a document-wide recomputation and a fresh cache.
	for _, entry := range changedProcedureEntries(catalog, request.PreviousCache, request.Changes) {
		if strings.HasPrefix(entry.Identity.Kind, "property_") {
			diagnostics := a.diagnosticsFullContext(ctx, doc)
			if ctx.Err() != nil {
				return DiagnosticResult{}
			}
			return DiagnosticResult{Diagnostics: diagnostics, Cache: buildDiagnosticCache(catalog, diagnostics)}
		}
	}
	return DiagnosticResult{Diagnostics: a.fastDiagnosticsContext(ctx, doc, catalog, request), Cache: request.PreviousCache}
}

func (a Analyzer) fastDiagnosticsContext(ctx context.Context, doc Document, catalog ProcedureCatalog, request DiagnosticRequest) []Diagnostic {
	changed := changedProcedureEntries(catalog, request.PreviousCache, request.Changes)
	changedKeys := make(map[ProcedureIdentity]bool, len(changed))
	out := make([]Diagnostic, 0)
	modulePreamble := ""
	modulePreambleEnd := 0
	if len(catalog.Entries) > 0 {
		modulePreambleEnd = catalog.Entries[0].StartByte
		modulePreamble = doc.Source[:modulePreambleEnd]
	}
	for _, entry := range changed {
		if ctx.Err() != nil {
			return nil
		}
		changedKeys[entry.Identity] = true
		fragment := doc
		fragment.Snapshot = nil
		fragment.Source = fastDiagnosticFragmentSource(doc.Source, modulePreamble, modulePreambleEnd, entry)
		fragmentCatalog := procedureCatalogForDocumentMode(fragment, false)
		if len(fragmentCatalog.Entries) != 1 {
			continue
		}
		fragmentEntry := fragmentCatalog.Entries[0]
		for _, diagnostic := range a.diagnosticsFullContext(ctx, fragment) {
			if fastDiagnosticForProcedure(diagnostic, fragmentEntry) {
				out = append(out, rebaseDiagnostic(diagnostic, fragmentEntry.Range.Start, entry.Range.Start))
			}
		}
	}
	if previous := request.PreviousCache; previous != nil && catalog.ReuseSafe && previous.Catalog.ReuseSafe && previous.Catalog.ModuleContextHash == catalog.ModuleContextHash && previous.Catalog.ConditionalHash == catalog.ConditionalHash {
		current := make(map[ProcedureIdentity]ProcedureCatalogEntry, len(catalog.Entries))
		for _, entry := range catalog.Entries {
			current[entry.Identity] = entry
		}
		for identity, cached := range previous.Procedures {
			entry, ok := current[identity]
			if !ok || changedKeys[identity] || entry.SourceHash != cached.Entry.SourceHash {
				continue
			}
			for _, diagnostic := range cached.Diagnostics {
				if !interproceduralDiagnostic(diagnostic.Code) {
					out = append(out, rebaseDiagnostic(diagnostic, cached.Entry.Range.Start, entry.Range.Start))
				}
			}
		}
	}
	out = applyDocumentInlineSuppressions(a.RootDir, doc, out)
	sort.SliceStable(out, func(i, j int) bool {
		if out[i].Range.Start.Line != out[j].Range.Start.Line {
			return out[i].Range.Start.Line < out[j].Range.Start.Line
		}
		if out[i].Range.Start.Character != out[j].Range.Start.Character {
			return out[i].Range.Start.Character < out[j].Range.Start.Character
		}
		return out[i].Code < out[j].Code
	})
	return out
}

type conditionalFragmentFrame struct {
	opening      string
	openingStart int
	current      string
	currentStart int
}

// fastDiagnosticFragmentSource preserves declarations before the first
// procedure and the conditional-compilation branch that encloses entry. The
// matching synthetic #End If directives make the isolated fragment a complete
// conditional-compilation unit for the parser.
func fastDiagnosticFragmentSource(source, modulePreamble string, modulePreambleEnd int, entry ProcedureCatalogEntry) string {
	var stack []conditionalFragmentFrame
	offset := 0
	for _, line := range strings.SplitAfter(source, "\n") {
		if offset >= entry.StartByte {
			break
		}
		trimmed := strings.TrimSpace(strings.TrimRight(line, "\r\n"))
		lower := strings.ToLower(trimmed)
		switch {
		case strings.HasPrefix(lower, "#if "):
			stack = append(stack, conditionalFragmentFrame{opening: line, openingStart: offset, current: line, currentStart: offset})
		case strings.HasPrefix(lower, "#elseif ") || lower == "#else":
			if len(stack) > 0 {
				stack[len(stack)-1].current = line
				stack[len(stack)-1].currentStart = offset
			}
		case lower == "#end if" || lower == "#endif":
			if len(stack) > 0 {
				stack = stack[:len(stack)-1]
			}
		}
		offset += len(line)
	}

	var prefix strings.Builder
	for _, frame := range stack {
		if frame.openingStart >= modulePreambleEnd {
			prefix.WriteString(frame.opening)
		}
		if frame.currentStart >= modulePreambleEnd && frame.currentStart != frame.openingStart {
			prefix.WriteString(frame.current)
		}
	}
	fragment := modulePreamble + prefix.String() + source[entry.StartByte:entry.EndByte]
	if len(stack) == 0 {
		return fragment
	}
	lineEnding := "\n"
	if strings.Contains(source, "\r\n") {
		lineEnding = "\r\n"
	}
	if !strings.HasSuffix(fragment, "\n") && !strings.HasSuffix(fragment, "\r") {
		fragment += lineEnding
	}
	for range stack {
		fragment += "#End If" + lineEnding
	}
	return fragment
}

func procedureCatalogForDocument(doc Document) ProcedureCatalog {
	if snapshot := analysisSnapshotForDocument(doc); snapshot != nil {
		snapshot.procedureCatalogOnce.Do(func() {
			snapshot.procedureCatalog = procedureCatalogForDocumentMode(doc, true)
		})
		return snapshot.procedureCatalog
	}
	return procedureCatalogForDocumentMode(doc, true)
}

func procedureCatalogForDocumentMode(doc Document, checkRecovery bool) ProcedureCatalog {
	procedures, _ := procedureIndexForLines(documentLines(doc))
	catalog := ProcedureCatalog{ReuseSafe: true}
	ordinals := make(map[string]int)
	conditionalHasher := sha256.New()
	lastByte := 0
	for _, procedure := range procedures {
		start := byteOffsetForDocumentPosition(doc, procedure.Range.Start)
		end := byteOffsetForDocumentPosition(doc, procedure.Range.End)
		if end < len(doc.Source) {
			_, end = rawLineBounds(doc.Source, end)
		}
		if start < lastByte || end < start {
			catalog.ReuseSafe = false
			continue
		}
		declaration := firstProcedureDeclarationLine(doc.Source[start:end])
		declarationStart := start
		for declarationStart < end && (doc.Source[declarationStart] == ' ' || doc.Source[declarationStart] == '\t') {
			declarationStart++
		}
		lineStart := strings.LastIndexByte(doc.Source[:declarationStart], '\n') + 1
		kind := procedureDeclarationKind(declaration)
		key := strings.ToLower(strings.TrimSpace(procedure.Name)) + "|" + kind
		ordinal := ordinals[key]
		ordinals[key]++
		catalog.Entries = append(catalog.Entries, ProcedureCatalogEntry{
			Identity:      ProcedureIdentity{CanonicalName: strings.ToLower(strings.TrimSpace(procedure.Name)), Kind: kind, Ordinal: ordinal},
			SignatureHash: sha256.Sum256([]byte(strings.ToLower(strings.TrimSpace(declaration)))),
			SourceHash:    sha256.Sum256([]byte(doc.Source[start:end])), Range: procedure.Range,
			StartByte: start, DeclarationStartByte: declarationStart, StartColumn: declarationStart - lineStart + 1, EndByte: end,
		})
		lastByte = end
	}
	moduleHasher := sha256.New()
	_, _ = moduleHasher.Write([]byte(strings.ToLower(strings.TrimSpace(doc.ModuleKind))))
	_, _ = moduleHasher.Write([]byte{'\n'})
	_, _ = moduleHasher.Write([]byte(strings.ToLower(strings.TrimSpace(moduleNameForDocument(doc)))))
	_, _ = moduleHasher.Write([]byte{'\n'})
	_, _ = moduleHasher.Write([]byte(strings.ToLower(strings.TrimSpace(doc.Path))))
	_, _ = moduleHasher.Write([]byte{'\n'})
	moduleSource := maskProcedureSet(doc.Source, catalog, nil)
	for _, line := range strings.Split(strings.ReplaceAll(strings.ReplaceAll(moduleSource, "\r\n", "\n"), "\r", "\n"), "\n") {
		trimmed := strings.TrimSpace(line)
		if trimmed == "" || strings.HasPrefix(trimmed, "'") {
			continue
		}
		_, _ = moduleHasher.Write([]byte(strings.ToLower(trimmed)))
		_, _ = moduleHasher.Write([]byte{'\n'})
	}
	for _, line := range documentLines(doc) {
		trimmed := strings.TrimSpace(line)
		if strings.HasPrefix(trimmed, "#") {
			_, _ = conditionalHasher.Write([]byte(strings.ToLower(trimmed)))
			_, _ = conditionalHasher.Write([]byte{'\n'})
		}
	}
	copy(catalog.ModuleContextHash[:], moduleHasher.Sum(nil))
	copy(catalog.ConditionalHash[:], conditionalHasher.Sum(nil))
	if checkRecovery {
		if snapshot := analysisSnapshotForDocument(doc); snapshot != nil {
			if parsed, err := snapshot.ParsedDocument(); err == nil {
				_ = parsed.Read(func(view ast.ParsedView) error {
					if view.HasError || view.HasMissing {
						catalog.ReuseSafe = false
					}
					return nil
				})
			}
		}
	}
	return catalog
}

func firstProcedureDeclarationLine(source string) string {
	lines := strings.Split(strings.ReplaceAll(strings.ReplaceAll(source, "\r\n", "\n"), "\r", "\n"), "\n")
	var declaration strings.Builder
	for _, line := range lines {
		trimmed := strings.TrimSpace(line)
		if trimmed == "" {
			continue
		}
		declaration.WriteString(trimmed)
		if !strings.HasSuffix(trimmed, "_") {
			break
		}
	}
	return declaration.String()
}

func procedureDeclarationKind(declaration string) string {
	lower := strings.ToLower(declaration)
	for _, kind := range []string{"property get", "property let", "property set", "function", "sub"} {
		if strings.Contains(lower, kind+" ") {
			return strings.ReplaceAll(kind, " ", "_")
		}
	}
	return "procedure"
}

func buildDiagnosticCache(catalog ProcedureCatalog, diagnostics []Diagnostic) *DiagnosticCache {
	cache := &DiagnosticCache{Catalog: catalog, Procedures: make(map[ProcedureIdentity]cachedProcedureDiagnostics)}
	for _, entry := range catalog.Entries {
		cache.Procedures[entry.Identity] = cachedProcedureDiagnostics{Entry: entry}
	}
	for _, diagnostic := range diagnostics {
		index := sort.Search(len(catalog.Entries), func(i int) bool {
			return comparePosition(catalog.Entries[i].Range.Start, diagnostic.Range.Start) > 0
		}) - 1
		if index < 0 || !rangeContainedBy(diagnostic.Range, catalog.Entries[index].Range) {
			continue
		}
		entry := catalog.Entries[index]
		anchored := cache.Procedures[entry.Identity]
		anchored.Diagnostics = append(anchored.Diagnostics, diagnostic)
		cache.Procedures[entry.Identity] = anchored
	}
	return cache
}

// NewDiagnosticCache anchors a completed Full publication to the document's
// current procedure catalog. It lets lifecycle schedulers retain their legacy
// Full diagnostic callback while enabling later Fast reuse.
func NewDiagnosticCache(doc Document, diagnostics []Diagnostic) *DiagnosticCache {
	return buildDiagnosticCache(procedureCatalogForDocument(doc), diagnostics)
}

func changedProcedureEntries(catalog ProcedureCatalog, previous *DiagnosticCache, changes ProcedureChangeSet) []ProcedureCatalogEntry {
	if previous == nil {
		if len(changes.Ranges) > 0 {
			var selected []ProcedureCatalogEntry
			for _, entry := range catalog.Entries {
				for _, changedRange := range changes.Ranges {
					if rangesIntersect(entry.Range, changedRange) {
						selected = append(selected, entry)
						break
					}
				}
			}
			if len(selected) > 0 {
				return selected
			}
		}
		return append([]ProcedureCatalogEntry(nil), catalog.Entries...)
	}
	if !catalog.ReuseSafe || !previous.Catalog.ReuseSafe || catalog.ModuleContextHash != previous.Catalog.ModuleContextHash || catalog.ConditionalHash != previous.Catalog.ConditionalHash {
		return append([]ProcedureCatalogEntry(nil), catalog.Entries...)
	}
	old := make(map[ProcedureIdentity]ProcedureCatalogEntry, len(previous.Catalog.Entries))
	for _, entry := range previous.Catalog.Entries {
		old[entry.Identity] = entry
	}
	var changed []ProcedureCatalogEntry
	for _, entry := range catalog.Entries {
		if prior, ok := old[entry.Identity]; !ok || prior.SourceHash != entry.SourceHash || prior.SignatureHash != entry.SignatureHash {
			changed = append(changed, entry)
		}
	}
	return changed
}

func fastDiagnosticForProcedure(diagnostic Diagnostic, entry ProcedureCatalogEntry) bool {
	if interproceduralDiagnostic(diagnostic.Code) {
		return false
	}
	if metadata, ok := staticrules.Lookup(diagnostic.Code); ok && metadata.Scope == staticrules.ScopeProcedureLocal && rangeContainedBy(diagnostic.Range, entry.Range) {
		return true
	}
	switch diagnostic.Code {
	case "VB008", "VB009", "VB013", "VB014", "VB032":
		return diagnostic.Range.Start.Line >= entry.Range.Start.Line && diagnostic.Range.Start.Line <= entry.Range.End.Line
	default:
		return false
	}
}

func interproceduralDiagnostic(code string) bool {
	switch strings.ToUpper(code) {
	case "VBA206", "VBA218", "VBA237", "VB052", "VB053", "VB054":
		return true
	default:
		return false
	}
}

func rebaseDiagnostic(diagnostic Diagnostic, oldStart, newStart Position) Diagnostic {
	diagnostic.Range.Start = rebasePosition(diagnostic.Range.Start, oldStart, newStart)
	diagnostic.Range.End = rebasePosition(diagnostic.Range.End, oldStart, newStart)
	return diagnostic
}

func rebasePosition(position, oldStart, newStart Position) Position {
	lineDelta := position.Line - oldStart.Line
	character := position.Character
	if lineDelta == 0 {
		character = newStart.Character + position.Character - oldStart.Character
	}
	return Position{Line: newStart.Line + lineDelta, Character: max(0, character)}
}

func rangeContainedBy(inner, outer Range) bool {
	return comparePosition(inner.Start, outer.Start) >= 0 && comparePosition(inner.End, outer.End) <= 0
}

func rangesIntersect(left, right Range) bool {
	return comparePosition(left.End, right.Start) >= 0 && comparePosition(right.End, left.Start) >= 0
}

func (identity ProcedureIdentity) String() string {
	return fmt.Sprintf("%s:%s:%d", identity.Kind, identity.CanonicalName, identity.Ordinal)
}
