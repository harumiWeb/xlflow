package intel

import (
	"context"
	"errors"
	"reflect"
	"sync/atomic"
	"testing"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func buildSnapshotArtifactsTestInput(t *testing.T, path, source string) (*vbaast.ParsedDocument, procedureir.DocumentIR, vbacfg.Document) {
	t.Helper()
	parsed, err := vbaast.ParseDocument(path, []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	ir, err := procedureir.BuildParsed(procedureir.BuildOptions{Path: path, ModuleKind: "standard"}, parsed)
	if err != nil {
		parsed.Close()
		t.Fatal(err)
	}
	cfg, err := vbacfg.BuildDocumentContext(context.Background(), ir)
	if err != nil {
		parsed.Close()
		t.Fatal(err)
	}
	return parsed, ir, cfg
}

func TestNewAnalysisSnapshotWithArtifactsSeedsCompletedCachesDefensively(t *testing.T) {
	const source = "Sub Main()\n  Dim value As Long\n  value = 1\nEnd Sub\n"
	path := "Main.bas"
	parsed, ir, _ := buildSnapshotArtifactsTestInput(t, path, source)
	defaultRange := vbaast.Range{StartLine: 2, EndLine: 2, StartByte: 14, EndByte: 15}
	boundsRange := vbaast.Range{StartLine: 2, EndLine: 2, StartByte: 16, EndByte: 17}
	lowerRange := vbaast.Range{StartLine: 2, EndLine: 2, StartByte: 18, EndByte: 19}
	upperRange := vbaast.Range{StartLine: 2, EndLine: 2, StartByte: 20, EndByte: 21}
	ir.Procedures[0].Symbol.Parameters = []procedureir.Parameter{{
		DefaultRange: &defaultRange,
		BoundsRange:  &boundsRange,
		ArrayBounds:  []procedureir.ArrayBound{{LowerRange: &lowerRange, UpperRange: &upperRange}},
	}}
	cfg := vbacfg.BuildDocument(ir)
	wantIR := procedureir.Clone(ir)
	wantCFG := vbacfg.CloneDocument(cfg)
	snapshot := NewAnalysisSnapshotWithArtifacts(
		Document{Path: path, Source: source, ModuleKind: "standard", Version: 1},
		parsed,
		AnalysisArtifacts{ProcedureIR: ir, ControlFlow: cfg},
	)
	defer snapshot.Retire()

	var irLoads, cfgLoads atomic.Int32
	gotIR, hit, err := snapshot.ProcedureIR(func() (procedureir.DocumentIR, error) {
		irLoads.Add(1)
		return procedureir.DocumentIR{}, errors.New("seeded IR loader invoked")
	})
	if err != nil || !hit || !reflect.DeepEqual(gotIR, wantIR) {
		t.Fatalf("seeded procedure IR = (equal=%v, hit=%v, err=%v)", reflect.DeepEqual(gotIR, wantIR), hit, err)
	}
	gotCFG, hit, err := snapshot.ControlFlowGraphs(func() (vbacfg.Document, error) {
		cfgLoads.Add(1)
		return vbacfg.Document{}, errors.New("seeded CFG loader invoked")
	})
	if err != nil || !hit || !reflect.DeepEqual(gotCFG, wantCFG) {
		t.Fatalf("seeded control flow = (equal=%v, hit=%v, err=%v)", reflect.DeepEqual(gotCFG, wantCFG), hit, err)
	}
	if irLoads.Load() != 0 || cfgLoads.Load() != 0 {
		t.Fatalf("seeded loaders called = (IR %d, CFG %d)", irLoads.Load(), cfgLoads.Load())
	}

	// Mutating the preparation values after construction must not affect the
	// snapshot-owned caches.
	ir.Procedures[0].Symbol.Name = "mutated input"
	ir.Procedures[0].Statements[0].Text = "mutated input"
	ir.Procedures[0].Symbol.Parameters[0].DefaultRange.StartByte++
	ir.Procedures[0].Symbol.Parameters[0].BoundsRange.EndByte++
	ir.Procedures[0].Symbol.Parameters[0].ArrayBounds[0].LowerRange.StartByte++
	ir.Procedures[0].Symbol.Parameters[0].ArrayBounds[0].UpperRange.EndByte++
	cfg.Graphs[0].Procedure.Name = "mutated input"
	cfg.Graphs[0].Procedure.Parameters[0].DefaultRange.StartByte++
	cfg.Graphs[0].Procedure.Parameters[0].BoundsRange.EndByte++
	cfg.Graphs[0].Procedure.Parameters[0].ArrayBounds[0].LowerRange.StartByte++
	cfg.Graphs[0].Procedure.Parameters[0].ArrayBounds[0].UpperRange.EndByte++
	cfg.Graphs[0].Blocks[0].Kind = vbacfg.BlockUnknownExit
	gotIR, _, err = snapshot.ProcedureIR(nil)
	if err != nil || !reflect.DeepEqual(gotIR, wantIR) {
		t.Fatalf("input mutation leaked into procedure IR = (equal=%v, err=%v)", reflect.DeepEqual(gotIR, wantIR), err)
	}
	gotCFG, _, err = snapshot.ControlFlowGraphs(nil)
	if err != nil || !reflect.DeepEqual(gotCFG, wantCFG) {
		t.Fatalf("input mutation leaked into control flow = (equal=%v, err=%v)", reflect.DeepEqual(gotCFG, wantCFG), err)
	}

	// Mutating one defensive getter result must not affect the next result.
	gotIR.Procedures[0].Symbol.Parameters = append(gotIR.Procedures[0].Symbol.Parameters, procedureir.Parameter{Name: "mutated getter"})
	gotIR.Procedures[0].Statements[0].Text = "mutated getter"
	gotIR.Procedures[0].Symbol.Parameters[0].DefaultRange.StartByte++
	gotCFG.Graphs[0].Procedure.Name = "mutated getter"
	gotCFG.Graphs[0].Procedure.Parameters[0].DefaultRange.StartByte++
	gotCFG.Graphs[0].Blocks[0].Kind = vbacfg.BlockUnknownExit
	againIR, _, _ := snapshot.ProcedureIR(nil)
	againCFG, _, _ := snapshot.ControlFlowGraphs(nil)
	if !reflect.DeepEqual(againIR, wantIR) || !reflect.DeepEqual(againCFG, wantCFG) {
		t.Fatal("mutating seeded getter result changed the snapshot cache")
	}
}

func TestNewAnalysisSnapshotWithArtifactsCancellationPrecedesCacheHit(t *testing.T) {
	const source = "Sub Main()\nEnd Sub\n"
	parsed, ir, cfg := buildSnapshotArtifactsTestInput(t, "Main.bas", source)
	snapshot := NewAnalysisSnapshotWithArtifacts(
		Document{Path: "Main.bas", Source: source, ModuleKind: "standard", Version: 1},
		parsed,
		AnalysisArtifacts{ProcedureIR: ir, ControlFlow: cfg},
	)
	defer snapshot.Retire()

	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	var loads atomic.Int32
	_, hit, err := snapshot.ProcedureIRContext(ctx, func(context.Context) (procedureir.DocumentIR, error) {
		loads.Add(1)
		return procedureir.DocumentIR{}, nil
	})
	if !errors.Is(err, context.Canceled) || hit || loads.Load() != 0 {
		t.Fatalf("pre-canceled seeded IR = (loads=%d, hit=%v, err=%v)", loads.Load(), hit, err)
	}
	_, hit, err = snapshot.ControlFlowGraphsContext(ctx, func(context.Context) (vbacfg.Document, error) {
		loads.Add(1)
		return vbacfg.Document{}, nil
	})
	if !errors.Is(err, context.Canceled) || hit || loads.Load() != 0 {
		t.Fatalf("pre-canceled seeded CFG = (loads=%d, hit=%v, err=%v)", loads.Load(), hit, err)
	}
}

func TestNewAnalysisSnapshotWithArtifactsSeedsSuccessorFragmentReuse(t *testing.T) {
	const oldSource = "Sub A()\n  Dim value As Long\n  value = 1\nEnd Sub\nSub B()\n  Dim other As Long\n  other = 2\nEnd Sub\n"
	const newSource = "Sub A()\n  Dim value As Long\n  value = 10\nEnd Sub\nSub B()\n  Dim other As Long\n  other = 2\nEnd Sub\n"
	oldPath := "Module1.bas"
	oldParsed, oldIR, oldCFG := buildSnapshotArtifactsTestInput(t, oldPath, oldSource)
	oldSnapshot := NewAnalysisSnapshotWithArtifacts(
		Document{Path: oldPath, Source: oldSource, ModuleKind: "standard", Version: 1},
		oldParsed,
		AnalysisArtifacts{ProcedureIR: oldIR, ControlFlow: oldCFG},
	)
	newParsed, err := vbaast.ParseDocument(oldPath, []byte(newSource))
	if err != nil {
		oldSnapshot.Retire()
		t.Fatal(err)
	}
	newSnapshot := NewSuccessorAnalysisSnapshotWithParsedDocument(
		Document{Path: oldPath, Source: newSource, ModuleKind: "standard", Version: 2},
		newParsed,
		oldSnapshot,
	)
	oldSnapshot.Retire()
	defer newSnapshot.Retire()

	newDoc := newSnapshot.Document()
	newIR, err := procedureIRForDocumentContext(context.Background(), newDoc, ".", newParsed)
	if err != nil {
		t.Fatal(err)
	}
	newCFG, err := controlFlowForDocumentContext(context.Background(), newDoc, newIR)
	if err != nil {
		t.Fatal(err)
	}
	stats := newSnapshot.ProcedureArtifactStats()
	if stats.IRBuild != 1 || stats.IRReuse != 1 || stats.CFGBuild != 1 || stats.CFGReuse != 1 {
		t.Fatalf("successor artifact stats = %+v, want one build and one reuse for both IR and CFG", stats)
	}
	if len(newIR.Procedures) != 2 || len(newCFG.Graphs) != 2 {
		t.Fatalf("successor artifacts = (IR %d procedures, CFG %d graphs)", len(newIR.Procedures), len(newCFG.Graphs))
	}
}
