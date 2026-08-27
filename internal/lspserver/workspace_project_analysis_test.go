package lspserver

import (
	"context"
	"crypto/sha256"
	"errors"
	"io"
	"log"
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

func TestIncrementalProjectDependencyViewReusesUnchangedProcedures(t *testing.T) {
	root := t.TempDir()
	a := filepath.Join(root, "A.bas")
	b := filepath.Join(root, "B.bas")
	c := filepath.Join(root, "C.bas")

	before := projectTestSnapshot(
		projectTestProcedure(a, "A.Run", "B.Work", b, 1, "Run"),
		projectTestProcedure(b, "B.Work", "C.Leaf", c, 1, "Work"),
		projectTestProcedure(c, "C.Leaf", "", "", 1, "old"),
	)
	for i := range before.Documents {
		before.Documents[i].Version = "v1"
	}
	view := buildProjectDependencyViewWithPerformanceClass(before, nil, "background")
	view.revision = 1

	after := projectTestSnapshot(
		projectTestProcedure(a, "A.Run", "B.Work", b, 1, "Run"),
		projectTestProcedure(b, "B.Work", "C.Leaf", c, 1, "Work"),
		projectTestProcedure(c, "C.Leaf", "", "", 1, "new"),
	)
	for i := range after.Documents {
		after.Documents[i].Version = "v1"
	}
	after.Documents[2].Version = "v2"
	recorder := newPerformanceRecorder(true, log.New(io.Discard, "", 0))
	impacted := updateProjectDependencyView(&view, after, recorder, "background")
	if got := recorder.counterTotal(performanceCounterProcedureFingerprintBuilds); got != 1 {
		t.Fatalf("fingerprint builds = %d, want only changed procedure", got)
	}
	if got := recorder.counterTotal(performanceCounterProcedureFingerprintReuses); got != 2 {
		t.Fatalf("fingerprint reuses = %d, want unchanged procedures", got)
	}
	if got := recorder.counterTotal(performanceCounterDependencyNodesUpdated); got != 1 {
		t.Fatalf("dependency nodes updated = %d, want changed procedure", got)
	}
	if !containsPath(impacted, a) || !containsPath(impacted, b) || !containsPath(impacted, c) {
		t.Fatalf("incremental impact paths = %#v, want caller closure", impacted)
	}
}

func TestIncrementalProjectDependencyViewPreservesUnchangedEdges(t *testing.T) {
	root := t.TempDir()
	caller := filepath.Join(root, "Caller.bas")
	callee := filepath.Join(root, "Callee.bas")
	before := projectTestSnapshot(
		projectTestProcedure(caller, "Caller.Run", "Callee.Work", callee, 1, "old"),
		projectTestProcedure(callee, "Callee.Work", "", "", 1, "Work"),
	)
	for i := range before.Documents {
		before.Documents[i].Version = "v1"
	}
	view := buildProjectDependencyViewWithPerformanceClass(before, nil, "background")
	after := projectTestSnapshot(
		projectTestProcedure(caller, "Caller.Run", "Callee.Work", callee, 1, "new"),
		projectTestProcedure(callee, "Callee.Work", "", "", 1, "Work"),
	)
	for i := range after.Documents {
		after.Documents[i].Version = "v1"
	}
	after.Documents[0].Version = "v2"
	recorder := newPerformanceRecorder(true, log.New(io.Discard, "", 0))
	updateProjectDependencyView(&view, after, recorder, "background")
	if got := recorder.counterTotal(performanceCounterDependencyEdgesUpdated); got != 0 {
		t.Fatalf("unchanged dependency edges updated = %d, want 0", got)
	}
}

func TestIncrementalProjectDependencyViewCountsDeletedProcedureEdges(t *testing.T) {
	path := filepath.Join(t.TempDir(), "Main.bas")
	first := projectTestProcedure(path, "Main.First", "", "", 1, "First").procedure
	second := projectTestProcedure(path, "Main.Second", "Main.First", path, 1, "Second").procedure
	beforeIR := procedureir.DocumentIR{Path: path, ModuleName: "Main", ModuleKind: "standard", Procedures: []procedureir.ProcedureIR{first, second}}
	before := intel.ProjectAnalysisSnapshot{Complete: true, Documents: []intel.ProjectAnalysisDocument{{
		IR: beforeIR, Version: "v1", CFG: vbacfg.BuildDocument(beforeIR), ProcedureCatalog: projectTestProcedureCatalog(beforeIR.Procedures),
	}}}
	view := buildProjectDependencyViewWithPerformanceClass(before, nil, "background")
	afterIR := procedureir.DocumentIR{Path: path, ModuleName: "Main", ModuleKind: "standard", Procedures: []procedureir.ProcedureIR{first}}
	after := intel.ProjectAnalysisSnapshot{Complete: true, Documents: []intel.ProjectAnalysisDocument{{
		IR: afterIR, Version: "v2", CFG: vbacfg.BuildDocument(afterIR), ProcedureCatalog: projectTestProcedureCatalog(afterIR.Procedures),
	}}}
	recorder := newPerformanceRecorder(true, log.New(io.Discard, "", 0))
	updateProjectDependencyView(&view, after, recorder, "background")
	if got := recorder.counterTotal(performanceCounterDependencyNodesUpdated); got != 1 {
		t.Fatalf("deleted dependency nodes updated = %d, want 1", got)
	}
	if got := recorder.counterTotal(performanceCounterDependencyEdgesUpdated); got != 1 {
		t.Fatalf("deleted dependency edges updated = %d, want 1", got)
	}
}

func TestIncrementalProjectDependencyViewDoesNotReuseUnsafeCatalog(t *testing.T) {
	path := filepath.Join(t.TempDir(), "Main.bas")
	procedure := projectTestProcedure(path, "Main.Run", "", "", 1, "Run").procedure
	ir := procedureir.DocumentIR{Path: path, ModuleName: "Main", ModuleKind: "standard", Procedures: []procedureir.ProcedureIR{procedure}}
	unsafeCatalog := projectTestProcedureCatalog(ir.Procedures)
	unsafeCatalog.ReuseSafe = false
	snapshot := intel.ProjectAnalysisSnapshot{Complete: true, Documents: []intel.ProjectAnalysisDocument{{
		IR: ir, Version: "v1", CFG: vbacfg.BuildDocument(ir), ProcedureCatalog: unsafeCatalog,
	}}}
	view := buildProjectDependencyViewWithPerformanceClass(snapshot, nil, "background")
	recorder := newPerformanceRecorder(true, log.New(io.Discard, "", 0))
	updateProjectDependencyView(&view, snapshot, recorder, "background")
	if got := recorder.counterTotal(performanceCounterProcedureFingerprintReuses); got != 0 {
		t.Fatalf("unsafe catalog fingerprint reuses = %d, want 0", got)
	}
	if got := recorder.counterTotal(performanceCounterProcedureFingerprintBuilds); got != 1 {
		t.Fatalf("unsafe catalog fingerprint builds = %d, want 1", got)
	}
}

func TestIncrementalProjectDependencyViewIncludesDeletedFiles(t *testing.T) {
	root := t.TempDir()
	caller := filepath.Join(root, "Caller.bas")
	callee := filepath.Join(root, "Callee.bas")
	before := projectTestSnapshot(
		projectTestProcedure(caller, "Caller.Run", "Callee.Work", callee, 1, "Run"),
		projectTestProcedure(callee, "Callee.Work", "", "", 1, "Work"),
	)
	for i := range before.Documents {
		before.Documents[i].Version = "v1"
	}
	view := buildProjectDependencyViewWithPerformanceClass(before, nil, "background")
	view.revision = 1
	after := projectTestSnapshot(projectTestProcedure(caller, "Caller.Run", "", "", 1, "Run"))
	after.Documents[0].Version = "v2"

	impacted := updateProjectDependencyView(&view, after, nil, "background")
	if !containsPath(impacted, caller) || !containsPath(impacted, callee) {
		t.Fatalf("deleted callee impact paths = %#v, want caller and deleted callee", impacted)
	}
}

func TestProjectImpactPathsIncludesResolvedAccessChanges(t *testing.T) {
	path := filepath.Join(t.TempDir(), "Main.bas")
	makeSnapshot := func(target string) intel.ProjectAnalysisSnapshot {
		procedure := procedureir.ProcedureIR{
			Symbol: procedureir.ProcedureSymbol{Name: "Run", QualifiedName: "Main.Run", Kind: procedureir.ProcedureSub},
			Accesses: []procedureir.VariableAccess{{
				Name: "Value", Resolution: procedureir.SymbolResolution{Status: procedureir.ResolutionMatched, Candidates: []procedureir.Candidate{{QualifiedName: target, Kind: "enum_member"}}},
			}},
		}
		ir := procedureir.DocumentIR{Path: path, ModuleName: "Main", ModuleKind: "standard", Procedures: []procedureir.ProcedureIR{procedure}}
		return intel.ProjectAnalysisSnapshot{Complete: true, Documents: []intel.ProjectAnalysisDocument{{IR: ir, CFG: vbacfg.BuildDocument(ir)}}}
	}
	impacted := projectImpactPaths(makeSnapshot("Enums.First"), makeSnapshot("Enums.Second"))
	if !containsPath(impacted, path) {
		t.Fatalf("access resolution impact paths = %#v, want %q", impacted, path)
	}
}

func TestProjectImpactPathsIncludesDeclarationModuleChanges(t *testing.T) {
	root := t.TempDir()
	callerPath := filepath.Join(root, "Caller.bas")
	declarationPath := filepath.Join(root, "Enums.bas")
	makeSnapshot := func(moduleHash byte) intel.ProjectAnalysisSnapshot {
		callerProcedure := procedureir.ProcedureIR{
			Symbol: procedureir.ProcedureSymbol{Name: "Run", QualifiedName: "Caller.Run", Kind: procedureir.ProcedureSub},
			Accesses: []procedureir.VariableAccess{{
				Name: "First", Resolution: procedureir.SymbolResolution{Status: procedureir.ResolutionMatched, Scope: procedureir.ScopeProject, Candidates: []procedureir.Candidate{{QualifiedName: "Enums.First", Kind: "enum_member", File: declarationPath, Line: 1}}},
			}},
		}
		callerIR := procedureir.DocumentIR{Path: callerPath, ModuleName: "Caller", ModuleKind: "standard", Procedures: []procedureir.ProcedureIR{callerProcedure}}
		declarationIR := procedureir.DocumentIR{Path: declarationPath, ModuleName: "Enums", ModuleKind: "standard", Declarations: []procedureir.Declaration{{Name: "First", Kind: "enum_member", Range: vbaast.Range{StartLine: 1}}}}
		var contextHash [sha256.Size]byte
		contextHash[0] = moduleHash
		return intel.ProjectAnalysisSnapshot{Complete: true, Documents: []intel.ProjectAnalysisDocument{
			{IR: callerIR, CFG: vbacfg.BuildDocument(callerIR)},
			{IR: declarationIR, ProcedureCatalog: intel.ProcedureCatalog{ModuleContextHash: contextHash, ConditionalHash: sha256.Sum256([]byte("declaration-conditional")), ReuseSafe: true}, CFG: vbacfg.BuildDocument(declarationIR)},
		}}
	}
	impacted := projectImpactPaths(makeSnapshot(1), makeSnapshot(2))
	if !containsPath(impacted, callerPath) || !containsPath(impacted, declarationPath) {
		t.Fatalf("declaration impact paths = %#v, want caller and declaration module", impacted)
	}
}

func TestProjectImpactPathsFailsOpenForUncertainResolution(t *testing.T) {
	root := t.TempDir()
	callerPath := filepath.Join(root, "Caller.bas")
	calleePath := filepath.Join(root, "Callee.bas")
	makeSnapshot := func(text string) intel.ProjectAnalysisSnapshot {
		caller := projectTestProcedure(callerPath, "Caller.Run", "", "", 1, "Run").procedure
		caller.Calls = []procedureir.CallSite{{Resolution: procedureir.CallResolution{Status: procedureir.ResolutionDynamic}}}
		callee := projectTestProcedure(calleePath, "Callee.Work", "", "", 1, text).procedure
		return projectTestSnapshot(
			projectTestProcedureInput{file: callerPath, procedure: caller},
			projectTestProcedureInput{file: calleePath, procedure: callee},
		)
	}
	impacted := projectImpactPaths(makeSnapshot("old"), makeSnapshot("new"))
	if !containsPath(impacted, callerPath) {
		t.Fatalf("uncertain resolution impact paths = %#v, want %q", impacted, callerPath)
	}
}

func TestIncrementalProjectDependencyViewRefreshesForDeclarationOnlyModule(t *testing.T) {
	root := t.TempDir()
	callerPath := filepath.Join(root, "Caller.bas")
	declarationPath := filepath.Join(root, "Enums.bas")
	makeSnapshot := func(moduleHash byte, target string, declarationVersion string) intel.ProjectAnalysisSnapshot {
		callerProcedure := procedureir.ProcedureIR{
			Symbol: procedureir.ProcedureSymbol{Name: "Run", QualifiedName: "Caller.Run", Kind: procedureir.ProcedureSub},
			Accesses: []procedureir.VariableAccess{{
				Name: "Value", Resolution: procedureir.SymbolResolution{Status: procedureir.ResolutionMatched, Candidates: []procedureir.Candidate{{QualifiedName: target, Kind: "enum_member"}}},
			}},
		}
		callerIR := procedureir.DocumentIR{Path: callerPath, ModuleName: "Caller", ModuleKind: "standard", Procedures: []procedureir.ProcedureIR{callerProcedure}}
		declarationIR := procedureir.DocumentIR{Path: declarationPath, ModuleName: "Enums", ModuleKind: "standard"}
		var contextHash [32]byte
		contextHash[0] = moduleHash
		return intel.ProjectAnalysisSnapshot{Complete: true, Documents: []intel.ProjectAnalysisDocument{
			{IR: callerIR, Version: "v1", CFG: vbacfg.BuildDocument(callerIR)},
			{IR: declarationIR, Version: declarationVersion, ProcedureCatalog: intel.ProcedureCatalog{ModuleContextHash: contextHash, ConditionalHash: sha256.Sum256([]byte("declaration-conditional")), ReuseSafe: true}, CFG: vbacfg.BuildDocument(declarationIR)},
		}}
	}
	before := makeSnapshot(1, "Enums.First", "v1")
	view := buildProjectDependencyViewWithPerformanceClass(before, nil, "background")
	view.revision = 1
	after := makeSnapshot(2, "Enums.Second", "v2")
	impacted := updateProjectDependencyView(&view, after, nil, "background")
	if !containsPath(impacted, callerPath) {
		t.Fatalf("declaration-only module impact paths = %#v, want %q", impacted, callerPath)
	}
}

func BenchmarkProjectDependencyIncrementalBodyEdit(b *testing.B) {
	const procedureCount = 2000
	before := largeProjectDependencyBenchmarkSnapshot(procedureCount, false)
	after := largeProjectDependencyBenchmarkSnapshot(procedureCount, true)
	benchmarkProjectDependencyUpdate(b, before, after)
}

func BenchmarkProjectDependencyIncrementalSignatureEdit(b *testing.B) {
	before := largeProjectDependencyBenchmarkSnapshot(2000, false)
	after := largeProjectDependencyBenchmarkSnapshot(2000, false)
	after.Documents[0].Version = "v2"
	after.Documents[0].ProcedureCatalog.Entries[1000].SignatureHash = sha256.Sum256([]byte("Private Sub"))
	benchmarkProjectDependencyUpdate(b, before, after)
}

func BenchmarkProjectDependencyIncrementalHighFanIn(b *testing.B) {
	before := highFanInProjectDependencyBenchmarkSnapshot(2000, false)
	after := highFanInProjectDependencyBenchmarkSnapshot(2000, true)
	benchmarkProjectDependencyUpdate(b, before, after)
}

func BenchmarkProjectDependencyIncrementalDeepCallChain(b *testing.B) {
	before := deepCallChainProjectDependencyBenchmarkSnapshot(1000, false)
	after := deepCallChainProjectDependencyBenchmarkSnapshot(1000, true)
	benchmarkProjectDependencyUpdate(b, before, after)
}

func BenchmarkProjectDependencyIncrementalCallDeletion(b *testing.B) {
	before := callChangeProjectDependencyBenchmarkSnapshot("Callee.Work", false)
	after := callChangeProjectDependencyBenchmarkSnapshot("", true)
	benchmarkProjectDependencyUpdate(b, before, after)
}

func BenchmarkProjectDependencyIncrementalCallTargetChange(b *testing.B) {
	before := callChangeProjectDependencyBenchmarkSnapshot("Callee.First", false)
	after := callChangeProjectDependencyBenchmarkSnapshot("Callee.Second", false)
	after.Documents[0].Version = "v2"
	benchmarkProjectDependencyUpdate(b, before, after)
}

// This synthetic 2k-procedure case is the default ROneCOne-scale dependency
// update benchmark. The optional real-file ROneCOne fixture remains covered by
// the existing opt-in LSP benchmark in issue491_benchmark_test.go.
func BenchmarkProjectDependencyIncrementalROneCOneScale(b *testing.B) {
	before := largeProjectDependencyBenchmarkSnapshot(2000, false)
	after := largeProjectDependencyBenchmarkSnapshot(2000, true)
	benchmarkProjectDependencyUpdate(b, before, after)
}

func benchmarkProjectDependencyUpdate(b *testing.B, before, after intel.ProjectAnalysisSnapshot) {
	b.Helper()
	base := buildProjectDependencyViewWithPerformanceClass(before, nil, "background")
	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		b.StopTimer()
		view := cloneProjectDependencyView(base)
		b.StartTimer()
		updateProjectDependencyView(&view, after, nil, "background")
	}
}

func highFanInProjectDependencyBenchmarkSnapshot(count int, edited bool) intel.ProjectAnalysisSnapshot {
	path := "FanIn.bas"
	procedures := make([]procedureir.ProcedureIR, count)
	procedures[0] = procedureir.ProcedureIR{
		Symbol:     procedureir.ProcedureSymbol{Name: "Callee", QualifiedName: "FanIn.Callee", Kind: procedureir.ProcedureSub},
		Statements: []procedureir.Statement{{ID: 1, Kind: procedureir.StatementCall, Text: map[bool]string{false: "same", true: "changed"}[edited]}},
	}
	for index := 1; index < count; index++ {
		name := "Caller" + decimalString(index)
		procedures[index] = procedureir.ProcedureIR{
			Symbol:     procedureir.ProcedureSymbol{Name: name, QualifiedName: "FanIn." + name, Kind: procedureir.ProcedureSub},
			Statements: []procedureir.Statement{{ID: index + 1, Kind: procedureir.StatementCall, Text: "Callee"}},
			Calls: []procedureir.CallSite{{
				ID: index + 1, Caller: procedureir.ProcedureRef{Name: name, QualifiedName: "FanIn." + name, Kind: procedureir.ProcedureSub},
				Callee:     procedureir.Callee{Text: "Callee", BaseName: "Callee"},
				Resolution: procedureir.CallResolution{Status: procedureir.ResolutionMatched, Candidates: []procedureir.Candidate{{QualifiedName: "FanIn.Callee", Kind: "sub", File: path, Line: 1}}},
			}},
		}
	}
	return benchmarkProjectDependencySnapshot(path, "FanIn", procedures, edited)
}

func deepCallChainProjectDependencyBenchmarkSnapshot(count int, edited bool) intel.ProjectAnalysisSnapshot {
	path := "Chain.bas"
	procedures := make([]procedureir.ProcedureIR, count)
	for index := range procedures {
		name := "Proc" + decimalString(index)
		procedure := procedureir.ProcedureIR{Symbol: procedureir.ProcedureSymbol{Name: name, QualifiedName: "Chain." + name, Kind: procedureir.ProcedureSub}, Statements: []procedureir.Statement{{ID: index + 1, Kind: procedureir.StatementCall, Text: "same"}}}
		if index < count-1 {
			target := "Proc" + decimalString(index+1)
			procedure.Calls = []procedureir.CallSite{{ID: index + 1, Caller: procedureir.ProcedureRef{Name: name, QualifiedName: "Chain." + name, Kind: procedureir.ProcedureSub}, Callee: procedureir.Callee{Text: target, BaseName: target}, Resolution: procedureir.CallResolution{Status: procedureir.ResolutionMatched, Candidates: []procedureir.Candidate{{QualifiedName: "Chain." + target, Kind: "sub", File: path, Line: 1}}}}}
		}
		if edited && index == count-1 {
			procedure.Statements[0].Text = "changed"
		}
		procedures[index] = procedure
	}
	return benchmarkProjectDependencySnapshot(path, "Chain", procedures, edited)
}

func callChangeProjectDependencyBenchmarkSnapshot(target string, deleted bool) intel.ProjectAnalysisSnapshot {
	path := "Caller.bas"
	caller := procedureir.ProcedureIR{Symbol: procedureir.ProcedureSymbol{Name: "Run", QualifiedName: "Caller.Run", Kind: procedureir.ProcedureSub}, Statements: []procedureir.Statement{{ID: 1, Kind: procedureir.StatementCall, Text: target}}}
	if !deleted {
		caller.Calls = []procedureir.CallSite{{ID: 1, Caller: procedureir.ProcedureRef{Name: "Run", QualifiedName: "Caller.Run", Kind: procedureir.ProcedureSub}, Callee: procedureir.Callee{Text: target, BaseName: target}, Resolution: procedureir.CallResolution{Status: procedureir.ResolutionMatched, Candidates: []procedureir.Candidate{{QualifiedName: target, Kind: "sub", File: "Callee.bas", Line: 1}}}}}
	}
	callee := procedureir.ProcedureIR{Symbol: procedureir.ProcedureSymbol{Name: "First", QualifiedName: "Callee.First", Kind: procedureir.ProcedureSub}}
	second := procedureir.ProcedureIR{Symbol: procedureir.ProcedureSymbol{Name: "Second", QualifiedName: "Callee.Second", Kind: procedureir.ProcedureSub}}
	return benchmarkProjectDependencySnapshot(path, "Caller", []procedureir.ProcedureIR{caller, callee, second}, deleted)
}

func benchmarkProjectDependencySnapshot(path, module string, procedures []procedureir.ProcedureIR, edited bool) intel.ProjectAnalysisSnapshot {
	ir := procedureir.DocumentIR{Path: path, ModuleName: module, ModuleKind: "standard", Procedures: procedures}
	version := "v1"
	if edited {
		version = "v2"
	}
	return intel.ProjectAnalysisSnapshot{Revision: 1, Complete: true, Documents: []intel.ProjectAnalysisDocument{{IR: ir, Version: version, ProcedureCatalog: projectTestProcedureCatalog(procedures), CFG: vbacfg.BuildDocument(ir)}}}
}

func largeProjectDependencyBenchmarkSnapshot(count int, edited bool) intel.ProjectAnalysisSnapshot {
	path := "Large.bas"
	procedures := make([]procedureir.ProcedureIR, count)
	entries := make([]intel.ProcedureCatalogEntry, count)
	for i := range procedures {
		text := "same"
		if edited && i == count/2 {
			text = "changed"
		}
		procedures[i] = procedureir.ProcedureIR{
			Symbol: procedureir.ProcedureSymbol{
				Name: qualifiedBenchmarkProcedureName(i), QualifiedName: "Large." + qualifiedBenchmarkProcedureName(i),
				Kind: procedureir.ProcedureSub, DeclarationRange: vbaast.Range{StartLine: i * 3},
			},
			Statements: []procedureir.Statement{{ID: i + 1, Kind: procedureir.StatementCall, Text: text}},
		}
		entries[i] = intel.ProcedureCatalogEntry{
			Identity:   intel.ProcedureIdentity{CanonicalName: "proc" + decimalString(i), Kind: "sub", Ordinal: i},
			SourceHash: sha256.Sum256([]byte(text)), SignatureHash: sha256.Sum256([]byte("Public Sub")),
		}
	}
	return intel.ProjectAnalysisSnapshot{Revision: 1, Complete: true, Documents: []intel.ProjectAnalysisDocument{{
		IR: procedureir.DocumentIR{Path: path, ModuleName: "Large", ModuleKind: "standard", Procedures: procedures}, Version: map[bool]string{false: "v1", true: "v2"}[edited],
		ProcedureCatalog: intel.ProcedureCatalog{Entries: entries, ModuleContextHash: sha256.Sum256([]byte("large-module")), ConditionalHash: sha256.Sum256([]byte("large-conditional")), ReuseSafe: true},
	}}}
}

func qualifiedBenchmarkProcedureName(index int) string {
	return "Proc" + decimalString(index)
}

func cloneProjectDependencyView(view projectDependencyView) projectDependencyView {
	clone := newProjectDependencyView()
	clone.revision = view.revision
	clone.lookup = cloneProjectProcedureLookup(view.lookup)
	for file, state := range view.files {
		state.procedureKeys = append([]string(nil), state.procedureKeys...)
		clone.files[file] = state
	}
	for key, state := range view.procedures {
		state.callees = cloneProjectCallers(state.callees)
		clone.procedures[key] = state
	}
	for callee, callers := range view.reverse {
		clone.reverse[callee] = cloneProjectCallers(callers)
	}
	return clone
}

func cloneProjectProcedureLookup(lookup projectProcedureLookup) projectProcedureLookup {
	clone := projectProcedureLookup{byQualified: make(map[string][]projectProcedureLookupEntry, len(lookup.byQualified))}
	for qualified, entries := range lookup.byQualified {
		clone.byQualified[qualified] = append([]projectProcedureLookupEntry(nil), entries...)
	}
	return clone
}

func TestCloneProjectDependencyViewPreservesLookup(t *testing.T) {
	base := buildProjectDependencyViewWithPerformanceClass(callChangeProjectDependencyBenchmarkSnapshot("Callee.First", false), nil, "background")
	clone := cloneProjectDependencyView(base)
	qualified := projectQualifiedKey("Callee.First", "sub")
	entries := clone.lookup.byQualified[qualified]
	if len(entries) != 1 {
		t.Fatalf("cloned lookup[%q] = %#v, want one entry", qualified, entries)
	}
	entries[0].key = "mutated"
	if base.lookup.byQualified[qualified][0].key == "mutated" {
		t.Fatal("cloned lookup shares entry storage with base view")
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
	resolution, ok := snapshot.Documents[0].Resolution.ResolvedCall(0, snapshot.Documents[0].IR.Procedures[0].Calls[0].ID)
	if !ok || resolution.Resolution.Status != procedureir.ResolutionMatched || len(resolution.Resolution.Candidates) != 1 {
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
	resolved, ok := snapshot.Documents[0].Resolution.ResolvedCall(0, snapshot.Documents[0].IR.Procedures[0].Calls[0].ID)
	if !ok {
		t.Fatal("call resolution was not attached to snapshot view")
	}
	candidates := resolved.Resolution.Candidates
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
	index.projectDependencies.revision = current.Revision + 1

	_, impacted := index.projectChange()
	if len(impacted) != 0 || index.projectDependencies.revision != current.Revision+1 {
		t.Fatalf("stale project change replaced baseline: impacted=%v baseline=%d", impacted, index.projectDependencies.revision)
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
	_, resolved, _, _, err := s.projectResolution(ctx, project, true)
	if err != nil {
		t.Fatal(err)
	}
	_, resolvedAgain, _, _, err := s.projectResolution(ctx, project, true)
	if err != nil {
		t.Fatal(err)
	}
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

func TestProjectResolutionCancellationLeavesCacheRetryable(t *testing.T) {
	project := projectTestSnapshot(projectTestProcedure(filepath.Join(t.TempDir(), "Main.bas"), "Main.Run", "", "", 1, "Run"))
	project.Revision = 21
	s := &Server{}
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	if _, _, _, _, err := s.projectResolution(ctx, project, true); !errors.Is(err, context.Canceled) {
		t.Fatalf("canceled project resolution error = %v, want context.Canceled", err)
	}
	if len(s.resolutionResolverCache.values) != 0 {
		t.Fatal("canceled project resolution was published to the cache")
	}
	if _, _, _, _, err := s.projectResolution(context.Background(), project, true); err != nil {
		t.Fatalf("project resolution retry error = %v", err)
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
		documents = append(documents, intel.ProjectAnalysisDocument{IR: ir, CFG: vbacfg.BuildDocument(ir), ProcedureCatalog: projectTestProcedureCatalog(ir.Procedures)})
	}
	return intel.ProjectAnalysisSnapshot{Complete: true, Documents: documents}
}

func projectTestProcedureCatalog(procedures []procedureir.ProcedureIR) intel.ProcedureCatalog {
	catalog := intel.ProcedureCatalog{ModuleContextHash: sha256.Sum256([]byte("test-module")), ConditionalHash: sha256.Sum256([]byte("test-conditional")), ReuseSafe: true, Entries: make([]intel.ProcedureCatalogEntry, 0, len(procedures))}
	for index, procedure := range procedures {
		hasher := sha256.New()
		writeFingerprintText(hasher, procedure.Symbol.QualifiedName, string(procedure.Symbol.Kind), procedure.Symbol.Name)
		for _, statement := range procedure.Statements {
			writeFingerprintText(hasher, string(statement.Kind), statement.Text, decimalString(statement.ID))
		}
		for _, call := range procedure.Calls {
			writeFingerprintText(hasher, string(call.Resolution.Status), call.Callee.Text, call.Callee.BaseName)
			for _, candidate := range call.Resolution.Candidates {
				writeFingerprintText(hasher, candidate.QualifiedName, candidate.Kind, candidate.File, decimalString(candidate.Line))
			}
		}
		var sourceHash [sha256.Size]byte
		copy(sourceHash[:], hasher.Sum(nil))
		catalog.Entries = append(catalog.Entries, intel.ProcedureCatalogEntry{
			Identity:      intel.ProcedureIdentity{CanonicalName: procedure.Symbol.Name, Kind: string(procedure.Symbol.Kind), Ordinal: index},
			SourceHash:    sourceHash,
			SignatureHash: sha256.Sum256([]byte(procedure.Symbol.QualifiedName)),
		})
	}
	return catalog
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
