package analyze

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestVBA202UsesProcedureIRObjectState(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private cachedSheet As Worksheet

Public Sub ForEachBody(ByVal wb As Workbook)
  Dim ws As Worksheet
  For Each ws In wb.Worksheets
    Debug.Print ws.Name
  Next ws
End Sub

Public Sub AfterForEach(ByVal wb As Workbook)
  Dim ws As Worksheet
  For Each ws In wb.Worksheets
    Debug.Print "visited"
  Next ws
  Debug.Print ws.Name
End Sub

Public Sub PartialAssignment(ByVal assignIt As Boolean)
  Dim ws As Worksheet
  If assignIt Then
    Set ws = ThisWorkbook.Worksheets(1)
  End If
  Debug.Print ws.Name
End Sub

Public Sub UninitializedRHS()
  Dim ws As Worksheet
  Dim target As Range
  Set target = ws.Range("A1")
End Sub

Public Sub ModuleStateIsUnknown()
  Debug.Print cachedSheet.Name
End Sub

Public Sub StaticStateIsPersistent()
  Static cachedRange As Range
  Debug.Print cachedRange.Address
End Sub

Public Sub TypeQualifierIsNotObjectUse()
  Dim Outlook As Outlook.Application
  Dim session As Outlook.NameSpace
End Sub

Public Sub UninitializedWithReceiver()
  Dim ws As Worksheet
  With ws
    Debug.Print .Name
  End With
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 3 {
		t.Fatalf("VBA202 findings = %+v, want after-loop, Set-RHS, and With-receiver findings", got)
	}
	wantLines := []int{16, 30, 49}
	for index, line := range wantLines {
		if got[index].Line != line {
			t.Fatalf("VBA202 finding %d = %+v, want line %d", index, got[index], line)
		}
	}
}
