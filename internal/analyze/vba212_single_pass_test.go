package analyze

import (
	"context"
	"errors"
	"strings"
	"testing"
	"time"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func TestVBA212ScannerTraversesRootOnceAndPreservesProcedureOwnership(t *testing.T) {
	source := []byte(`Option Explicit
Public Sub First()
  Dim one As Collection
  If one Is Nothing Or one.Count = 0 Then Exit Sub
End Sub

Public Function Second() As Boolean
  Dim two As Collection
  If Not two Is Nothing And two.Count > 0 Then Second = True
End Function

Public Property Get Third() As Boolean
  Dim three As Collection
  If three Is Nothing Or three.Count = 0 Then Third = True
End Property
`)
	doc, err := vbaast.ParseDocument("Guard.cls", source)
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	ir, err := procedureir.BuildParsed(procedureir.BuildOptions{}, doc)
	if err != nil {
		t.Fatal(err)
	}
	stats := &vba212ScanStats{}
	var findings []Finding
	err = doc.Read(func(view vbaast.ParsedView) error {
		file := parsedFile{Path: view.Path, Module: "Guard", Source: view.Source, Root: view.Root, Lines: normalizedSourceLines(string(view.Source)), IR: ir}
		findings, err = (Analyzer{}).nonShortCircuitObjectGuardDocumentFindings(context.Background(), file, sourceProceduresFromIR(ir), stats)
		return err
	})
	if err != nil {
		t.Fatal(err)
	}
	if stats.RootTraversals != 1 {
		t.Fatalf("root traversals = %d, want 1", stats.RootTraversals)
	}
	if stats.NodesVisited == 0 {
		t.Fatal("expected scanner to visit CST nodes")
	}
	if len(findings) != 3 {
		t.Fatalf("VBA212 findings = %+v, want three", findings)
	}
	want := []string{"First", "Second", "Third"}
	for i := range want {
		if findings[i].Procedure != want[i] {
			t.Fatalf("finding %d procedure = %q, want %q: %+v", i, findings[i].Procedure, want[i], findings)
		}
	}
}

func TestVBA212ScannerCancellationStopsAtCheckpoint(t *testing.T) {
	var source strings.Builder
	source.WriteString("Option Explicit\nPublic Sub Run()\n  Dim item As Collection\n")
	for i := 0; i < 600; i++ {
		source.WriteString("  Debug.Print item.Count\n")
	}
	source.WriteString("End Sub\n")
	doc, err := vbaast.ParseDocument("Cancel.bas", []byte(source.String()))
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	ir, err := procedureir.BuildParsed(procedureir.BuildOptions{}, doc)
	if err != nil {
		t.Fatal(err)
	}
	stats := &vba212ScanStats{}
	ctx := vba212CheckpointContext{stats: stats, cancelAt: 256}
	err = doc.Read(func(view vbaast.ParsedView) error {
		file := parsedFile{Path: view.Path, Module: "Cancel", Source: view.Source, Root: view.Root, Lines: normalizedSourceLines(string(view.Source)), IR: ir}
		_, scanErr := (Analyzer{}).nonShortCircuitObjectGuardDocumentFindings(ctx, file, sourceProceduresFromIR(ir), stats)
		return scanErr
	})
	if !errors.Is(err, context.Canceled) {
		t.Fatalf("scan error = %v, want context.Canceled", err)
	}
	if stats.NodesVisited != 256 {
		t.Fatalf("nodes visited after cancellation = %d, want checkpoint 256", stats.NodesVisited)
	}
}

type vba212CheckpointContext struct {
	stats    *vba212ScanStats
	cancelAt int
}

func (c vba212CheckpointContext) Deadline() (time.Time, bool) { return time.Time{}, false }
func (c vba212CheckpointContext) Done() <-chan struct{}       { return nil }
func (c vba212CheckpointContext) Value(any) any               { return nil }
func (c vba212CheckpointContext) Err() error {
	if c.stats.NodesVisited >= c.cancelAt {
		return context.Canceled
	}
	return nil
}
