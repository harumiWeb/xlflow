package intel

import (
	"context"
	"reflect"
	"testing"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func TestSuccessorReusesUnchangedProcedureIRAndCFG(t *testing.T) {
	oldSource := "Attribute VB_Name = \"Module1\"\nOption Explicit\nSub A()\n  Dim x As Long\n  x = 1\nEnd Sub\nSub B()\n  Dim y As Long\n  y = 2\nEnd Sub\n"
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

	newSource := "Attribute VB_Name = \"Module1\"\nOption Explicit\nSub A()\n  Dim x As Long\n  x = 10\nEnd Sub\nSub B()\n  Dim y As Long\n  y = 2\nEnd Sub\n"
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
