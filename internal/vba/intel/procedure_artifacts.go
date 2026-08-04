package intel

import (
	"context"
	"sort"
	"sync"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

type procedureArtifactKey struct {
	Identity      ProcedureIdentity
	SourceHash    [32]byte
	ModuleContext [32]byte
}

type procedureIRFragment struct {
	base           vbaast.Range
	procedure      procedureir.ProcedureIR
	typeReferences []procedureir.TypeReference
}

type procedureArtifactStore struct {
	mu  sync.RWMutex
	ir  map[procedureArtifactKey]procedureIRFragment
	cfg map[procedureArtifactKey]vbacfg.Graph
}

func newProcedureArtifactStore() *procedureArtifactStore {
	return &procedureArtifactStore{ir: make(map[procedureArtifactKey]procedureIRFragment), cfg: make(map[procedureArtifactKey]vbacfg.Graph)}
}

func (s *procedureArtifactStore) clone() *procedureArtifactStore {
	if s == nil {
		return newProcedureArtifactStore()
	}
	s.mu.RLock()
	defer s.mu.RUnlock()
	out := newProcedureArtifactStore()
	for key, fragment := range s.ir {
		out.ir[key] = fragment
	}
	for key, graph := range s.cfg {
		out.cfg[key] = graph
	}
	return out
}

func artifactKey(entry ProcedureCatalogEntry, catalog ProcedureCatalog) procedureArtifactKey {
	return procedureArtifactKey{Identity: entry.Identity, SourceHash: entry.SourceHash, ModuleContext: catalog.ModuleContextHash}
}

func astBase(entry ProcedureCatalogEntry) vbaast.Range {
	return vbaast.Range{StartLine: entry.Range.Start.Line + 1, StartColumn: 1, StartByte: entry.StartByte}
}

func (s *AnalysisSnapshot) seedProcedureArtifacts(ir procedureir.DocumentIR) {
	if s == nil || s.artifacts == nil || ir.Parse.HasError || ir.Parse.HasMissing {
		return
	}
	catalog := procedureCatalogForDocument(s.Document())
	if !catalog.ReuseSafe || len(catalog.Entries) != len(ir.Procedures) {
		return
	}
	s.artifacts.mu.Lock()
	defer s.artifacts.mu.Unlock()
	for i, entry := range catalog.Entries {
		procedure := ir.Procedures[i]
		fragment := procedureIRFragment{base: procedure.Symbol.DeclarationRange, procedure: procedureir.CloneProcedureIR(procedure)}
		for _, reference := range ir.TypeReferences {
			if reference.Caller != nil && reference.Range.StartByte >= procedure.Symbol.DeclarationRange.StartByte && reference.Range.EndByte <= procedure.Symbol.DeclarationRange.EndByte {
				fragment.typeReferences = append(fragment.typeReferences, reference)
			}
		}
		s.artifacts.ir[artifactKey(entry, catalog)] = fragment
	}
}

func (s *AnalysisSnapshot) incrementalProcedureIR(ctx context.Context, rootDir string) (procedureir.DocumentIR, bool, error) {
	if s == nil || s.artifacts == nil {
		return procedureir.DocumentIR{}, false, nil
	}
	doc := s.Document()
	catalog := procedureCatalogForDocument(doc)
	if !catalog.ReuseSafe || len(catalog.Entries) == 0 {
		return procedureir.DocumentIR{}, false, nil
	}
	hits := make(map[ProcedureIdentity]procedureIRFragment)
	misses := make(map[ProcedureIdentity]bool)
	s.artifacts.mu.RLock()
	for _, entry := range catalog.Entries {
		fragment, ok := s.artifacts.ir[artifactKey(entry, catalog)]
		if ok {
			hits[entry.Identity] = fragment
		} else {
			misses[entry.Identity] = true
		}
	}
	s.artifacts.mu.RUnlock()
	if len(hits) == 0 {
		return procedureir.DocumentIR{}, false, nil
	}
	moduleSource := maskProcedureSet(doc.Source, catalog, nil)
	moduleIR, err := procedureir.BuildSourceContext(ctx, procedureir.BuildOptions{RootDir: rootDir, Path: doc.Path, ModuleKind: doc.ModuleKind}, []byte(moduleSource))
	if err != nil || moduleIR.Parse.HasError || moduleIR.Parse.HasMissing {
		return procedureir.DocumentIR{}, false, err
	}
	var built procedureir.DocumentIR
	if len(misses) > 0 {
		built, err = procedureir.BuildSourceContext(ctx, procedureir.BuildOptions{RootDir: rootDir, Path: doc.Path, ModuleKind: doc.ModuleKind}, []byte(maskProcedureSet(doc.Source, catalog, misses)))
		if err != nil || built.Parse.HasError || built.Parse.HasMissing {
			return procedureir.DocumentIR{}, false, err
		}
	}
	builtByStart := make(map[int]procedureir.ProcedureIR)
	for _, procedure := range built.Procedures {
		builtByStart[procedure.Symbol.DeclarationRange.StartByte] = procedure
	}
	result := moduleIR
	result.Procedures = make([]procedureir.ProcedureIR, 0, len(catalog.Entries))
	result.TypeReferences = append([]procedureir.TypeReference(nil), moduleIR.TypeReferences...)
	for _, entry := range catalog.Entries {
		if err := ctx.Err(); err != nil {
			return procedureir.DocumentIR{}, false, err
		}
		if fragment, ok := hits[entry.Identity]; ok {
			procedure := procedureir.RebaseProcedure(fragment.procedure, fragment.base, astBase(entry))
			result.Procedures = append(result.Procedures, procedure)
			for _, reference := range fragment.typeReferences {
				result.TypeReferences = append(result.TypeReferences, procedureir.RebaseTypeReference(reference, fragment.base, astBase(entry)))
			}
			s.procedureIRReuseCount.Add(1)
			continue
		}
		procedure, ok := builtByStart[entry.StartByte]
		if !ok {
			return procedureir.DocumentIR{}, false, nil
		}
		result.Procedures = append(result.Procedures, procedure)
		for _, reference := range built.TypeReferences {
			if reference.Caller != nil && reference.Range.StartByte >= entry.StartByte && reference.Range.EndByte <= entry.EndByte {
				result.TypeReferences = append(result.TypeReferences, reference)
			}
		}
		s.procedureIRBuildCount.Add(1)
	}
	result.Declarations = append([]procedureir.Declaration(nil), moduleIR.Declarations...)
	declarations := make([]*procedureir.Declaration, 0, len(result.Declarations))
	for index := range result.Declarations {
		declarations = append(declarations, &result.Declarations[index])
	}
	for procedureIndex := range result.Procedures {
		for declarationIndex := range result.Procedures[procedureIndex].Declarations {
			declarations = append(declarations, &result.Procedures[procedureIndex].Declarations[declarationIndex])
		}
	}
	sort.SliceStable(declarations, func(i, j int) bool { return declarations[i].Range.StartByte < declarations[j].Range.StartByte })
	for index, declaration := range declarations {
		declaration.ID = index + 1
	}
	return result, true, nil
}

func maskProcedureSet(source string, catalog ProcedureCatalog, keep map[ProcedureIdentity]bool) string {
	masked := []byte(source)
	for _, entry := range catalog.Entries {
		if keep != nil && keep[entry.Identity] {
			continue
		}
		for i := entry.StartByte; i < entry.EndByte && i < len(masked); i++ {
			if masked[i] != '\r' && masked[i] != '\n' {
				masked[i] = ' '
			}
		}
	}
	return string(masked)
}

func (s *AnalysisSnapshot) seedCFGArtifacts(ir procedureir.DocumentIR, document vbacfg.Document) {
	if s == nil || s.artifacts == nil || len(ir.Procedures) != len(document.Graphs) {
		return
	}
	catalog := procedureCatalogForDocument(s.Document())
	if len(catalog.Entries) != len(document.Graphs) {
		return
	}
	s.artifacts.mu.Lock()
	defer s.artifacts.mu.Unlock()
	for i, entry := range catalog.Entries {
		s.artifacts.cfg[artifactKey(entry, catalog)] = vbacfg.Clone(document.Graphs[i])
	}
}

func (s *AnalysisSnapshot) incrementalCFG(ctx context.Context, ir procedureir.DocumentIR) (vbacfg.Document, bool, error) {
	if s == nil || s.artifacts == nil {
		return vbacfg.Document{}, false, nil
	}
	catalog := procedureCatalogForDocument(s.Document())
	if len(catalog.Entries) != len(ir.Procedures) {
		return vbacfg.Document{}, false, nil
	}
	result := vbacfg.Document{Path: ir.Path, Graphs: make([]vbacfg.Graph, 0, len(ir.Procedures))}
	s.artifacts.mu.RLock()
	defer s.artifacts.mu.RUnlock()
	for i, entry := range catalog.Entries {
		if err := ctx.Err(); err != nil {
			return vbacfg.Document{}, false, err
		}
		if graph, ok := s.artifacts.cfg[artifactKey(entry, catalog)]; ok {
			result.Graphs = append(result.Graphs, vbacfg.RebaseGraph(graph, graph.Procedure.DeclarationRange, ir.Procedures[i].Symbol.DeclarationRange))
			s.cfgReuseCount.Add(1)
			continue
		}
		graph, err := vbacfg.BuildContext(ctx, ir.Procedures[i])
		if err != nil {
			return vbacfg.Document{}, false, err
		}
		result.Graphs = append(result.Graphs, graph)
		s.cfgBuildCount.Add(1)
	}
	return result, true, nil
}

type ProcedureArtifactStats struct {
	IRBuild, IRReuse, CFGBuild, CFGReuse uint64
}

func (s *AnalysisSnapshot) ProcedureArtifactStats() ProcedureArtifactStats {
	if s == nil {
		return ProcedureArtifactStats{}
	}
	return ProcedureArtifactStats{IRBuild: s.procedureIRBuildCount.Load(), IRReuse: s.procedureIRReuseCount.Load(), CFGBuild: s.cfgBuildCount.Load(), CFGReuse: s.cfgReuseCount.Load()}
}
