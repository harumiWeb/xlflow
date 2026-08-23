package lspserver

import (
	"context"
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
)

func TestProjectImpactPathsUsesOldAndNewReverseGraphsTransitively(t *testing.T) {
	root := t.TempDir()
	a := filepath.Join(root, "A.bas")
	b := filepath.Join(root, "B.bas")
	c := filepath.Join(root, "C.bas")
	d := filepath.Join(root, "D.bas")
	before := projectTestSnapshot(
		projectTestProcedure(a, "A.Run", "B.Work", b, 1, "Run"),
		projectTestProcedure(b, "B.Work", "C.Leaf", c, 1, "Work"),
		projectTestProcedure(c, "C.Leaf", "", "", 1, "old"),
		projectTestProcedure(d, "D.Unrelated", "", "", 1, "same"),
	)
	after := projectTestSnapshot(
		projectTestProcedure(a, "A.Run", "", "", 1, "Run"),
		projectTestProcedure(b, "B.Work", "C.Leaf", c, 1, "Work"),
		projectTestProcedure(c, "C.Leaf", "", "", 1, "new"),
		projectTestProcedure(d, "D.Unrelated", "", "", 1, "same"),
	)

	got := projectImpactPaths(before, after)
	for _, want := range []string{a, b, c} {
		if !containsPath(got, want) {
			t.Fatalf("impact paths = %#v, want %q", got, want)
		}
	}
	if containsPath(got, d) {
		t.Fatalf("impact paths = %#v, unrelated path %q was invalidated", got, d)
	}
}

func TestWorkspaceProjectSnapshotHoldsResolvedIRAndDefensiveCFG(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "Main.bas")
	ir := procedureir.DocumentIR{Path: path, ModuleName: "Main", ModuleKind: "standard", Procedures: []procedureir.ProcedureIR{
		projectTestProcedure(path, "Main.Run", "Main.Work", path, 4, "Run").procedure,
		projectTestProcedure(path, "Main.Work", "", "", 4, "Work").procedure,
	}}
	graph := vbacfg.BuildDocument(ir)
	index := newWorkspaceAnalysisIndex(root, config.Default(), func(context.Context, symbols.SourceFile, []byte) (indexedFileAnalysis, error) {
		return indexedFileAnalysis{}, nil
	}, nil)
	if err := index.waitReady(); err != nil {
		t.Fatal(err)
	}
	doc := intel.Document{URI: "file:///Main.bas", Path: path, Source: "source", ModuleKind: "standard"}
	index.setOverlay(doc, indexedFileAnalysis{
		procedureIR: ir, controlFlow: graph, source: doc.Source,
		symbols: []intel.Symbol{
			{Name: "Run", Kind: "sub", Module: "Main", ModuleKind: "standard", File: path, Range: intel.Range{Start: intel.Position{Line: 0}}},
			{Name: "Work", Kind: "sub", Module: "Main", ModuleKind: "standard", File: path, Range: intel.Range{Start: intel.Position{Line: 3}}},
		},
	})

	snapshot := index.projectSnapshot()
	if !snapshot.Complete || len(snapshot.Documents) != 1 || len(snapshot.Documents[0].CFG.Graphs) != 2 {
		t.Fatalf("snapshot = %#v", snapshot)
	}
	if snapshot.Documents[0].Source != doc.Source {
		t.Fatalf("snapshot source = %q, want %q", snapshot.Documents[0].Source, doc.Source)
	}
	if snapshot.Revision == 0 {
		t.Fatal("published project snapshot revision was not advanced")
	}
	resolution := snapshot.Documents[0].IR.Procedures[0].Calls[0].Resolution
	if resolution.Status != procedureir.ResolutionMatched || len(resolution.Candidates) != 1 {
		t.Fatalf("resolution = %#v", resolution)
	}
	snapshot.Documents[0].IR.Procedures[0].Symbol.Name = "Mutated"
	snapshot.Documents[0].CFG.Graphs[0].Blocks = nil
	again := index.projectSnapshot()
	if again.Documents[0].IR.Procedures[0].Symbol.Name != "Run" || len(again.Documents[0].CFG.Graphs[0].Blocks) == 0 {
		t.Fatal("project snapshot shares IR or CFG storage")
	}

	index.beginOverlay(doc, 2)
	pending := index.projectSnapshot()
	if pending.Complete || pending.Revision <= snapshot.Revision {
		t.Fatalf("pending overlay snapshot = %#v, want incomplete newer revision", pending)
	}
	if !index.abandonOverlay(doc, 2) {
		t.Fatal("failed overlay was not abandoned")
	}
	abandoned := index.projectSnapshot()
	if abandoned.Complete || abandoned.Revision <= pending.Revision {
		t.Fatal("failed overlay produced a complete project snapshot")
	}
}

func TestWorkspaceProjectSnapshotResolverCandidatesUseDisplayPaths(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "src", "modules", "Main.bas")
	call := procedureir.CallSite{
		File: path, Module: "Main",
		Caller: procedureir.ProcedureRef{Name: "Run", Kind: procedureir.ProcedureSub, QualifiedName: "Main.Run"},
		Callee: procedureir.Callee{Text: "Work", BaseName: "Work"},
	}
	ir := procedureir.DocumentIR{
		Path: path, ModuleName: "Main", ModuleKind: "standard",
		Procedures: []procedureir.ProcedureIR{
			{Symbol: procedureir.ProcedureSymbol{Name: "Run", QualifiedName: "Main.Run", Kind: procedureir.ProcedureSub}, Calls: []procedureir.CallSite{call}},
			{Symbol: procedureir.ProcedureSymbol{Name: "Work", QualifiedName: "Main.Work", Kind: procedureir.ProcedureSub}},
		},
	}
	index := newWorkspaceAnalysisIndex(root, config.Default(), func(context.Context, symbols.SourceFile, []byte) (indexedFileAnalysis, error) {
		return indexedFileAnalysis{}, nil
	}, nil)
	if err := index.waitReady(); err != nil {
		t.Fatal(err)
	}
	index.setOverlay(intel.Document{Path: path, ModuleKind: "standard", Version: 1}, indexedFileAnalysis{
		path: path, moduleKind: "standard", procedureIR: ir,
		symbols: []intel.Symbol{
			{Name: "Run", Kind: "sub", Module: "Main", ModuleKind: "standard", File: path},
			{Name: "Work", Kind: "sub", Module: "Main", ModuleKind: "standard", File: path},
		},
	})
	snapshot := index.projectSnapshot()
	if !snapshot.Complete || len(snapshot.Documents) != 1 {
		t.Fatalf("snapshot = %#v", snapshot)
	}
	candidates := snapshot.Documents[0].IR.Procedures[0].Calls[0].Resolution.Candidates
	if len(candidates) != 1 || candidates[0].File != filepath.ToSlash(filepath.Join("src", "modules", "Main.bas")) {
		t.Fatalf("resolver candidates = %#v, want workspace display path", candidates)
	}
}

func TestProjectChangeRejectsStaleSnapshotBaseline(t *testing.T) {
	index := newWorkspaceAnalysisIndex(t.TempDir(), config.Default(), func(context.Context, symbols.SourceFile, []byte) (indexedFileAnalysis, error) {
		return indexedFileAnalysis{}, nil
	}, nil)
	if err := index.waitReady(); err != nil {
		t.Fatal(err)
	}
	current := index.projectSnapshot()
	index.lastProjectSnapshot = intel.ProjectAnalysisSnapshot{Revision: current.Revision + 1, Complete: true}

	_, impacted := index.projectChange()
	if len(impacted) != 0 || index.lastProjectSnapshot.Revision != current.Revision+1 {
		t.Fatalf("stale project change replaced baseline: impacted=%v baseline=%d", impacted, index.lastProjectSnapshot.Revision)
	}
}

func TestProjectEffectSummaryCachesByRevision(t *testing.T) {
	input := projectTestProcedure(filepath.Join(t.TempDir(), "Main.bas"), "Main.Run", "", "", 1, "Run")
	project := projectTestSnapshot(input)
	project.Revision = 7
	s := &Server{}
	first := s.projectEffectSummary(project)
	if len(first.All()) != 1 {
		t.Fatalf("first summary procedures = %d, want 1", len(first.All()))
	}
	project.Documents = nil
	again := s.projectEffectSummary(project)
	if len(again.All()) != 1 {
		t.Fatalf("cached summary procedures = %d, want 1", len(again.All()))
	}
}

func TestProjectCapabilityCachesRecordOneBuildPerRevision(t *testing.T) {
	input := projectTestProcedure(filepath.Join(t.TempDir(), "Main.bas"), "Main.Run", "", "", 1, "Run")
	project := projectTestSnapshot(input)
	project.Revision = 9
	s := &Server{}
	recorder := analysisstats.NewRecorder()
	ctx := analysisstats.WithRecorder(context.Background(), recorder)
	_, resolved, _ := s.projectResolution(ctx, project, true)
	_, resolvedAgain, _ := s.projectResolution(ctx, project, true)
	if len(resolved) != len(resolvedAgain) {
		t.Fatalf("resolution cache sizes = %d and %d", len(resolved), len(resolvedAgain))
	}
	s.projectEffectSummaryWithResolution(ctx, project, resolved, true)
	s.projectEffectSummaryWithResolution(ctx, project, resolved, true)
	_, counters := recorder.Totals()
	values := map[string]uint64{}
	for _, counter := range counters {
		values[counter.Name] = counter.Value
	}
	if values[analysisstats.CapabilityResolutionBuildsCounter] != 1 {
		t.Fatalf("resolution builds = %d, want 1", values[analysisstats.CapabilityResolutionBuildsCounter])
	}
	if values[analysisstats.CapabilityEffectsBuildsCounter] != 1 {
		t.Fatalf("effects builds = %d, want 1", values[analysisstats.CapabilityEffectsBuildsCounter])
	}
}

func TestWorkspaceProjectSnapshotIsIncompleteDuringCloseDiskRefresh(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "Main.bas")
	index := newWorkspaceAnalysisIndex(root, config.Default(), func(context.Context, symbols.SourceFile, []byte) (indexedFileAnalysis, error) {
		return indexedFileAnalysis{}, nil
	}, nil)
	if err := index.waitReady(); err != nil {
		t.Fatal(err)
	}
	doc := intel.Document{URI: pathToFileURI(path), Path: path, Source: "source", ModuleKind: "standard"}
	index.setOverlay(doc, indexedFileAnalysis{procedureIR: procedureir.DocumentIR{Path: path}})
	refresh := index.beginClearOverlay(path)
	if refresh == nil {
		t.Fatal("close refresh was not reserved")
	}
	if snapshot := index.projectSnapshot(); snapshot.Complete {
		t.Fatalf("close refresh snapshot = %#v, want incomplete", snapshot)
	}
	if _, restored, err := index.finishClearOverlay(refresh); err != nil || restored {
		t.Fatalf("finish missing close refresh = (restored=%v, err=%v)", restored, err)
	}
	if snapshot := index.projectSnapshot(); !snapshot.Complete {
		t.Fatalf("completed deletion snapshot = %#v, want complete", snapshot)
	}
}

type projectTestProcedureInput struct {
	file      string
	procedure procedureir.ProcedureIR
}

func projectTestSnapshot(procedures ...projectTestProcedureInput) intel.ProjectAnalysisSnapshot {
	documents := make([]intel.ProjectAnalysisDocument, 0, len(procedures))
	for _, input := range procedures {
		procedure := input.procedure
		ir := procedureir.DocumentIR{Path: input.file, ModuleName: moduleFromQualified(procedure.Symbol.QualifiedName), ModuleKind: "standard", Procedures: []procedureir.ProcedureIR{procedure}}
		documents = append(documents, intel.ProjectAnalysisDocument{IR: ir, CFG: vbacfg.BuildDocument(ir)})
	}
	return intel.ProjectAnalysisSnapshot{Complete: true, Documents: documents}
}

func projectTestProcedure(file, qualified, target, targetFile string, targetLine int, text string) projectTestProcedureInput {
	module := moduleFromQualified(qualified)
	name := qualified[len(module)+1:]
	procedure := procedureir.ProcedureIR{Symbol: procedureir.ProcedureSymbol{
		Name: name, QualifiedName: qualified, Kind: procedureir.ProcedureSub,
		DeclarationRange: vbaast.Range{StartLine: 1}, BodyRange: vbaast.Range{StartLine: 1, EndLine: 3},
	}, Statements: []procedureir.Statement{{ID: 1, Kind: procedureir.StatementCall, Text: text, Range: vbaast.Range{StartLine: 2}}}}
	if target != "" {
		procedure.Calls = []procedureir.CallSite{{
			ID: 1, Caller: procedureir.ProcedureRef{Name: name, Kind: procedureir.ProcedureSub, QualifiedName: qualified},
			Callee:      procedureir.Callee{Text: target, BaseName: target[len(moduleFromQualified(target))+1:]},
			StatementID: 1, Resolution: procedureir.CallResolution{Status: procedureir.ResolutionMatched, Candidates: []procedureir.Candidate{{QualifiedName: target, Kind: "sub", File: targetFile, Line: targetLine}}},
		}}
	}
	return projectTestProcedureInput{file: file, procedure: procedure}
}

func moduleFromQualified(value string) string {
	for i, r := range value {
		if r == '.' {
			return value[:i]
		}
	}
	return value
}

func containsPath(paths []string, want string) bool {
	for _, path := range paths {
		if symbolFileKey(path) == symbolFileKey(want) {
			return true
		}
	}
	return false
}
