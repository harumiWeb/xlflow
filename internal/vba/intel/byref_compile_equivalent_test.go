package intel

import (
	"context"
	"path/filepath"
	"reflect"
	"testing"
)

func TestCompileEquivalentDiagnosticsWithPrecomputedByRefDiagnostics(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{
		Path: filepath.Join(t.TempDir(), "Main.bas"),
		Source: `Option Explicit
Public Sub NeedsLong(ByRef value As Long)
End Sub
Public Sub Run()
    Dim text As String
    NeedsLong text
End Sub
`,
	}

	ctx := context.Background()
	want := analyzer.CompileEquivalentDiagnosticsContext(ctx, doc)
	byRef := analyzer.ByRefArgumentDiagnosticsContext(ctx, doc)
	got := analyzer.CompileEquivalentDiagnosticsContextWithByRefDiagnostics(ctx, doc, byRef)

	if !reflect.DeepEqual(got, want) {
		t.Fatalf("precomputed ByRef compile-equivalent diagnostics changed:\n got: %+v\nwant: %+v", got, want)
	}
	if diagnostics := diagnosticsByCode(got, "VBA228"); len(diagnostics) != 1 {
		t.Fatalf("precomputed ByRef diagnostics = %+v, want one VBA228", diagnostics)
	}
}
