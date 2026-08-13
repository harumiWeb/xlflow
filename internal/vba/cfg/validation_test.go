package cfg

import (
	"testing"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func validationCodes(t *testing.T, source string) []string {
	t.Helper()
	ir, err := procedureir.BuildSource(procedureir.BuildOptions{Path: "Module1.bas"}, []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	document, err := BuildDocumentContext(t.Context(), ir)
	if err != nil {
		t.Fatal(err)
	}
	diagnostics := ValidationDiagnostics(document)
	codes := make([]string, 0, len(diagnostics))
	for _, diagnostic := range diagnostics {
		codes = append(codes, diagnostic.Code)
	}
	return codes
}

func TestValidationDiagnosticsLabelsAndTransfers(t *testing.T) {
	codes := validationCodes(t, `Sub Main()
    GoTo MissingLabel
    On Error GoTo MissingHandler
    Resume MissingResume
Label1:
label1:
10:
10:
End Sub`)
	want := []string{"VB056", "VB056", "VB056", "VB055", "VB055"}
	if len(codes) != len(want) {
		t.Fatalf("codes = %#v, want %#v", codes, want)
	}
	for i := range want {
		if codes[i] != want[i] {
			t.Fatalf("codes = %#v, want %#v", codes, want)
		}
	}
}

func TestValidationDiagnosticsExcludeGoSubAndComputedTransfers(t *testing.T) {
	codes := validationCodes(t, `Sub Main()
    GoSub MissingSub
    On x GoTo MissingOne, MissingTwo
    Return
End Sub`)
	for _, code := range codes {
		if code == "VB055" || code == "VB056" {
			t.Fatalf("unsupported transfer produced label diagnostic %s: %#v", code, codes)
		}
	}
}

func TestValidationDiagnosticsNextVariablesAndCanonicalization(t *testing.T) {
	codes := validationCodes(t, `Sub Main()
    For i% = 1 To 2
        For Each [j$] In values
        Next J$, I%
    Next i%
    For i = 1 To 2
    Next wrong
End Sub`)
	if len(codes) != 1 || codes[0] != "VB057" {
		t.Fatalf("codes = %#v, want one VB057", codes)
	}
}

func TestValidationDiagnosticsNextExtraVariable(t *testing.T) {
	codes := validationCodes(t, `Sub Main()
    For i = 1 To 2
    Next i, extra
End Sub`)
	if len(codes) != 1 || codes[0] != "VB057" {
		t.Fatalf("codes = %#v, want one VB057", codes)
	}
}

func TestValidationDiagnosticsLoopAndProcedureExits(t *testing.T) {
	codes := validationCodes(t, `Sub Main()
    Exit Function
    Exit For
    Do
        For i = 1 To 2
            Exit Do
            Exit For
        Next
    Loop
    Exit Do
End Sub

Function F()
    Exit Function
    Exit Property
End Function

Property Get P()
    Exit Property
    Exit Function
End Property`)
	want := []string{"VB058", "VB058", "VB058", "VB058", "VB058"}
	if len(codes) != len(want) {
		t.Fatalf("codes = %#v, want %#v", codes, want)
	}
	for i := range want {
		if codes[i] != want[i] {
			t.Fatalf("codes = %#v, want %#v", codes, want)
		}
	}
}

func TestValidationDiagnosticsFailOpenOnRecovery(t *testing.T) {
	codes := validationCodes(t, `Sub Main()
    For i = 1 To
        Next wrong
    GoTo Missing
End Sub`)
	for _, code := range codes {
		if code == "VB056" || code == "VB057" || code == "VB058" {
			t.Fatalf("recovered source produced speculative code %s: %#v", code, codes)
		}
	}
}

func TestValidationDiagnosticsUnparsedNextFailsOpen(t *testing.T) {
	if !unparsedNextStatement(procedureir.Statement{Kind: procedureir.StatementCall, Text: "Next outer"}) {
		t.Fatal("expected parser-leftover Next call to be uncertain")
	}
	if unparsedNextStatement(procedureir.Statement{Kind: procedureir.StatementCall, Text: "NextValue outer"}) {
		t.Fatal("identifier beginning with Next must not be treated as a Next statement")
	}
}

func TestValidationFactsCloneAndRebasePreserveRanges(t *testing.T) {
	graph := Build(buildProcedure(t, `Sub Main()
Label:
label:
End Sub`))
	if len(graph.ValidationFacts) != 1 {
		t.Fatalf("validation facts = %#v", graph.ValidationFacts)
	}
	clone := Clone(graph)
	clone.ValidationFacts[0].Range.StartByte++
	if clone.ValidationFacts[0].Range.StartByte == graph.ValidationFacts[0].Range.StartByte {
		t.Fatal("clone shares validation fact storage")
	}
	oldBase := vbaast.Range{StartLine: 1, StartColumn: 1, StartByte: 0}
	newBase := vbaast.Range{StartLine: 4, StartColumn: 1, StartByte: 30}
	rebased := RebaseGraph(graph, oldBase, newBase)
	if rebased.ValidationFacts[0].Range.StartLine != graph.ValidationFacts[0].Range.StartLine+3 ||
		rebased.ValidationFacts[0].Range.StartByte != graph.ValidationFacts[0].Range.StartByte+30 {
		t.Fatalf("rebased validation fact = %#v, original = %#v", rebased.ValidationFacts[0], graph.ValidationFacts[0])
	}
}
