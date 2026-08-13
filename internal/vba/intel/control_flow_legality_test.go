package intel

import (
	"context"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestCompileEquivalentDiagnosticsProjectsCFGControlFlowLegality(t *testing.T) {
	doc := Document{Path: "Main.bas", Source: `Sub Main()
    GoTo Missing
    Label1:
    label1:
    For i = 1 To 2
    Next wrong
    Exit Function
End Sub`}
	diagnostics := (Analyzer{RootDir: ".", Config: config.Default()}).CompileEquivalentDiagnosticsContext(context.Background(), doc)
	seen := map[string]int{}
	for _, diagnostic := range diagnostics {
		if diagnostic.Code >= "VB055" && diagnostic.Code <= "VB058" {
			seen[diagnostic.Code]++
		}
	}
	for _, code := range []string{"VB055", "VB056", "VB057", "VB058"} {
		if seen[code] != 1 {
			t.Fatalf("%s count = %d, diagnostics = %+v", code, seen[code], diagnostics)
		}
	}
}

func TestCompileEquivalentDiagnosticsUsesUTF16CFGRange(t *testing.T) {
	doc := Document{Path: "Main.bas", Source: "Sub Main()\n    '日本語\n    GoTo Missing\nEnd Sub\n"}
	diagnostics := (Analyzer{RootDir: "."}).CompileEquivalentDiagnosticsContext(context.Background(), doc)
	for _, diagnostic := range diagnostics {
		if diagnostic.Code != "VB056" {
			continue
		}
		if diagnostic.Range.Start.Line != 2 || diagnostic.Range.Start.Character != 9 {
			t.Fatalf("VB056 start = %+v, want line 2 character 9", diagnostic.Range.Start)
		}
		if diagnostic.Range.End.Line != 2 || diagnostic.Range.End.Character != 16 {
			t.Fatalf("VB056 end = %+v, want line 2 character 16", diagnostic.Range.End)
		}
		return
	}
	t.Fatalf("VB056 missing: %+v", diagnostics)
}

func TestFullDiagnosticsReuseLintCFGProjectionWithoutDuplicates(t *testing.T) {
	doc := Document{Path: "Main.bas", Source: `Sub Main()
    GoTo Missing
    Label1:
    label1:
End Sub`}
	diagnostics := (Analyzer{RootDir: ".", Config: config.Default()}).DiagnosticsContext(context.Background(), doc)
	seen := map[string]int{}
	for _, diagnostic := range diagnostics {
		if diagnostic.Code >= "VB055" && diagnostic.Code <= "VB058" {
			seen[diagnostic.Code]++
		}
	}
	if seen["VB055"] != 1 || seen["VB056"] != 1 {
		t.Fatalf("CFG diagnostics = %+v", diagnostics)
	}
	for _, diagnostic := range diagnostics {
		if diagnostic.Code == "VB056" {
			if diagnostic.Range.Start.Line != 1 || diagnostic.Range.Start.Character != 9 ||
				diagnostic.Range.End.Line != 1 || diagnostic.Range.End.Character != 16 {
				t.Fatalf("full VB056 range = %+v, want target token range", diagnostic.Range)
			}
		}
	}
}

func TestFastDiagnosticsProjectsCFGControlFlowLegality(t *testing.T) {
	doc := Document{Path: "Main.bas", Source: "Sub Main()\n    GoTo Missing\nEnd Sub\n"}
	result := (Analyzer{RootDir: ".", Config: config.Default()}).DiagnosticsRequestContext(context.Background(), DiagnosticRequest{
		Document: doc, Mode: DiagnosticModeFast,
		Changes: ProcedureChangeSet{Ranges: []Range{{Start: Position{Line: 0, Character: 0}, End: Position{Line: 3, Character: 0}}}},
	})
	for _, diagnostic := range result.Diagnostics {
		if diagnostic.Code == "VB056" {
			return
		}
	}
	t.Fatalf("fast CFG diagnostics = %+v", result.Diagnostics)
}
