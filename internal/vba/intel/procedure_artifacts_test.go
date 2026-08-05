package intel

import (
	"context"
	"fmt"
	"reflect"
	"strings"
	"testing"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func TestSuccessorReusesUnchangedProcedureIRAndCFG(t *testing.T) {
	oldSource := "Attribute VB_Name = \"Module1\"\nOption Explicit\nSub A()\n  Dim x As Long\n  x = 1\nEnd Sub\n  Sub B()\n  Dim y As Long\n  y = 2\n  End Sub\n"
	oldDoc := Document{Path: "Module1.bas", Source: oldSource, ModuleKind: "standard", Version: 1}
	oldSnapshot := NewAnalysisSnapshot(oldDoc)
	oldDoc = oldSnapshot.Document()
	parsed, err := oldSnapshot.ParsedDocument()
	if err != nil {
		t.Fatal(err)
	}
	oldIR, err := procedureIRForDocumentContext(context.Background(), oldDoc, ".", parsed)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := controlFlowForDocumentContext(context.Background(), oldDoc, oldIR); err != nil {
		t.Fatal(err)
	}

	newSource := "Attribute VB_Name = \"Module1\"\nOption Explicit\nSub A()\n  Dim x As Long\n  x = 10\nEnd Sub\n  Sub B()\n  Dim y As Long\n  y = 2\n  End Sub\n"
	newDoc := Document{Path: oldDoc.Path, Source: newSource, ModuleKind: oldDoc.ModuleKind, Version: 2}
	newParsed, err := vbaast.ParseDocument(newDoc.Path, []byte(newSource))
	if err != nil {
		t.Fatal(err)
	}
	newSnapshot := NewSuccessorAnalysisSnapshotWithParsedDocument(newDoc, newParsed, oldSnapshot)
	oldSnapshot.Retire()
	defer newSnapshot.Retire()
	newDoc = newSnapshot.Document()
	newIR, err := procedureIRForDocumentContext(context.Background(), newDoc, ".", newParsed)
	if err != nil {
		t.Fatal(err)
	}
	newCFG, err := controlFlowForDocumentContext(context.Background(), newDoc, newIR)
	if err != nil {
		t.Fatal(err)
	}
	fullIR, err := procedureir.BuildSourceContext(context.Background(), procedureir.BuildOptions{RootDir: ".", Path: newDoc.Path, ModuleKind: newDoc.ModuleKind}, []byte(newSource))
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(procedureir.Clone(newIR), procedureir.Clone(fullIR)) {
		t.Fatalf("incremental IR differs from full rebuild:\nincremental=%+v\nfull=%+v", newIR, fullIR)
	}
	fullCFG, err := vbacfg.BuildDocumentContext(context.Background(), fullIR)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(newCFG, fullCFG) {
		t.Fatalf("incremental CFG differs from full rebuild")
	}
	stats := newSnapshot.ProcedureArtifactStats()
	if stats.IRBuild != 1 || stats.IRReuse != 1 {
		t.Fatalf("IR stats = %+v, want build=1 reuse=1", stats)
	}
	if stats.CFGBuild != 1 || stats.CFGReuse != 1 {
		t.Fatalf("CFG stats = %+v, want build=1 reuse=1", stats)
	}
}

func TestProcedureArtifactSeedingMatchesDeclarationsByStartByte(t *testing.T) {
	source := "Attribute VB_Name = \"Module1\"\nOption Explicit\nSub A()\nEnd Sub\nSub B()\nEnd Sub\n"
	doc := Document{Path: "Module1.bas", Source: source, ModuleKind: "standard", Version: 1}
	snapshot := NewAnalysisSnapshot(doc)
	defer snapshot.Retire()

	ir, err := procedureir.BuildSourceContext(context.Background(), procedureir.BuildOptions{RootDir: ".", Path: doc.Path, ModuleKind: doc.ModuleKind}, []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	cfg, err := vbacfg.BuildDocumentContext(context.Background(), ir)
	if err != nil {
		t.Fatal(err)
	}
	ir.Procedures[0], ir.Procedures[1] = ir.Procedures[1], ir.Procedures[0]
	cfg.Graphs[0], cfg.Graphs[1] = cfg.Graphs[1], cfg.Graphs[0]
	snapshot.seedProcedureArtifacts(ir)
	snapshot.seedCFGArtifacts(ir, cfg)

	catalog := procedureCatalogForDocument(snapshot.Document())
	for _, entry := range catalog.Entries {
		key := artifactKey(entry, catalog)
		fragment, ok := snapshot.artifacts.ir[key]
		if !ok || !strings.EqualFold(fragment.procedure.Symbol.Name, entry.Identity.CanonicalName) {
			t.Fatalf("IR artifact for %s = (%+v, %t)", entry.Identity, fragment.procedure.Symbol, ok)
		}
		graph, ok := snapshot.artifacts.cfg[key]
		if !ok || !strings.EqualFold(graph.Procedure.Name, entry.Identity.CanonicalName) {
			t.Fatalf("CFG artifact for %s = (%+v, %t)", entry.Identity, graph.Procedure, ok)
		}
	}
}

func TestCloseReopenSnapshotDoesNotInheritProcedureArtifacts(t *testing.T) {
	doc := Document{Path: "Module1.bas", Source: "Sub A()\nEnd Sub\n", ModuleKind: "standard", Version: 1}
	first := NewAnalysisSnapshot(doc)
	parsed, err := first.ParsedDocument()
	if err != nil {
		t.Fatal(err)
	}
	if _, err := procedureIRForDocumentContext(context.Background(), first.Document(), ".", parsed); err != nil {
		t.Fatal(err)
	}
	first.Retire()
	second := NewAnalysisSnapshot(doc)
	defer second.Retire()
	if len(second.artifacts.ir) != 0 || len(second.artifacts.cfg) != 0 {
		t.Fatal("fresh close/reopen snapshot inherited procedure artifacts")
	}
}

func TestProcedureArtifactStorePrunesObsoleteRevisions(t *testing.T) {
	newSource := func(value int) string {
		if value == 25 {
			return fmt.Sprintf("Attribute VB_Name = \"Module1\"\nOption Explicit\nSub A()\n  Dim x As Long\n  x = %d\nEnd Sub\n", value)
		}
		return fmt.Sprintf("Attribute VB_Name = \"Module1\"\nOption Explicit\nSub A()\n  Dim x As Long\n  x = %d\nEnd Sub\nSub B()\n  Dim y As Long\n  y = 2\nEnd Sub\n", value)
	}

	var previous *AnalysisSnapshot
	for revision := 1; revision <= 25; revision++ {
		source := newSource(revision)
		doc := Document{Path: "Module1.bas", Source: source, ModuleKind: "standard", Version: int32(revision)}
		parsed, err := vbaast.ParseDocument(doc.Path, []byte(source))
		if err != nil {
			t.Fatal(err)
		}
		current := NewSuccessorAnalysisSnapshotWithParsedDocument(doc, parsed, previous)
		if previous != nil {
			previous.Retire()
		}
		doc = current.Document()
		ir, err := procedureIRForDocumentContext(context.Background(), doc, ".", parsed)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := controlFlowForDocumentContext(context.Background(), doc, ir); err != nil {
			t.Fatal(err)
		}
		current.artifacts.mu.RLock()
		irCount, cfgCount := len(current.artifacts.ir), len(current.artifacts.cfg)
		current.artifacts.mu.RUnlock()
		want := 2
		if revision == 25 {
			want = 1
		}
		if irCount != want || cfgCount != want {
			current.Retire()
			t.Fatalf("revision %d artifact counts = (IR %d, CFG %d), want (%d, %d)", revision, irCount, cfgCount, want, want)
		}
		previous = current
	}
	previous.Retire()
}

func TestRecoveredSnapshotDoesNotReuseCFGFragments(t *testing.T) {
	validSource := "Sub A()\n  If True Then\n  End If\nEnd Sub\nSub B()\nEnd Sub\n"
	validDoc := Document{Path: "Module1.bas", Source: validSource, ModuleKind: "standard", Version: 1}
	validSnapshot := NewAnalysisSnapshot(validDoc)
	validDoc = validSnapshot.Document()
	parsed, err := validSnapshot.ParsedDocument()
	if err != nil {
		t.Fatal(err)
	}
	ir, err := procedureIRForDocumentContext(context.Background(), validDoc, ".", parsed)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := controlFlowForDocumentContext(context.Background(), validDoc, ir); err != nil {
		t.Fatal(err)
	}

	recoveredSource := "Public Function A(ByVal value As String\nEnd Function\nSub B()\nEnd Sub\n"
	recoveredDoc := Document{Path: validDoc.Path, Source: recoveredSource, ModuleKind: validDoc.ModuleKind, Version: 2}
	recoveredParsed, err := vbaast.ParseDocument(recoveredDoc.Path, []byte(recoveredSource))
	if err != nil {
		t.Fatal(err)
	}
	recoveredSnapshot := NewSuccessorAnalysisSnapshotWithParsedDocument(recoveredDoc, recoveredParsed, validSnapshot)
	validSnapshot.Retire()
	defer recoveredSnapshot.Retire()
	if procedureCatalogForDocument(recoveredSnapshot.Document()).ReuseSafe {
		t.Fatal("recovered procedure catalog unexpectedly allowed fragment reuse")
	}
	if _, reused, err := recoveredSnapshot.incrementalCFG(context.Background(), ir); err != nil || reused {
		t.Fatalf("recovered CFG reuse = (reused=%v, err=%v), want conservative fallback", reused, err)
	}
}
