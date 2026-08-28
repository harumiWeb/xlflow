package intel

import (
	"context"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
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

func TestBatchTypeInferenceDoesNotUseInteractiveScopeIndex(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Private Type tGUID
  Value As Long
End Type

#If VBA7 Then
Public Property Get hwnd() As LongPtr
#Else
Public Property Get hwnd() As Long
#End If
  Dim unrelated As Long
End Property

Private Function getIAccessible() As tGUID
  Dim Guid As tGUID
  getIAccessible = Guid
End Function

Private Function convertGUID(Guid As String) As tGUID
  Dim vArr: vArr = Split(Guid, "-")
End Function

Private Function Split(Expression As String, Optional Delimiter As String = " ") As String()
End Function
`
	doc := Document{Path: filepath.Join(t.TempDir(), "stdAcc.cls"), Source: source, ModuleKind: "class"}
	parsed, err := vbaast.ParseDocument(doc.Path, []byte(doc.Source))
	if err != nil {
		t.Fatal(err)
	}
	defer parsed.Close()
	snapshot := NewAnalysisSnapshotWithParsedDocument(doc, parsed)
	doc = snapshot.Document()
	defer snapshot.Retire()

	line := lineIndex(source, "Dim vArr:")
	if line < 0 {
		t.Fatal("test source lost convertGUID body")
	}
	inferred, ok := analyzer.inferWordTypeInfoAt(doc, "Guid", byteOffsetForDocumentPosition(doc, Position{Line: line, Character: 0}))
	if ok && strings.EqualFold(inferred.Type, "tGUID") {
		t.Fatalf("batch type inference selected an unrelated local declaration: %+v; interactive scope data must not override the full document index", inferred)
	}
}
