package analyze

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestVBA202Issue448TracksObjectStateAcrossCFGAndCalls(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private sharedSheet As Worksheet

Private Sub Init(ByRef target As Worksheet, ByVal assignIt As Boolean)
  If assignIt Then Set target = ThisWorkbook.Worksheets(1)
End Sub

Private Sub ClearTarget(ByRef target As Worksheet)
  Set target = Nothing
End Sub

Public Function MaybeSheet(ByVal returnIt As Boolean) As Worksheet
  If returnIt Then Set MaybeSheet = ThisWorkbook.Worksheets(1)
End Function

Public Sub Run(ByVal assignIt As Boolean)
  Dim ws As Worksheet
  Init ws, assignIt
  Debug.Print ws.Name
  Debug.Print sharedSheet.Name
End Sub

Public Sub ResetUse()
  Dim resetTarget As Worksheet
  Set resetTarget = ThisWorkbook.Worksheets(1)
  ClearTarget resetTarget
  Debug.Print resetTarget.Name
End Sub

Public Sub ReturnUse()
  Dim maybeTarget As Worksheet
  Set maybeTarget = MaybeSheet(False)
  Debug.Print maybeTarget.Name
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	wantLines := []int{19, 20, 27, 33}
	if len(got) != len(wantLines) {
		t.Fatalf("VBA202 findings = %+v, want lines %v", got, wantLines)
	}
	for index, line := range wantLines {
		if got[index].Line != line {
			t.Fatalf("VBA202 finding %d = %+v, want line %d", index, got[index], line)
		}
	}
}

func TestVBA202Issue448RecognizesAsNewAndConstructors(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim dict As New Scripting.Dictionary
  Dim created As Object
  Set created = CreateObject("Scripting.Dictionary")
  Debug.Print dict.Count
  Debug.Print created.Count
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("constructor-backed objects should not produce VBA202: %+v", got)
	}
}

func TestVBA202Issue448ChecksCollectionAndDictionaryDefaultItems(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Unsafe()
  Dim values As Collection
  Debug.Print values(1)
End Sub

Public Sub Safe()
  Dim values As Object
  Set values = CreateObject("Scripting.Dictionary")
  Debug.Print values("key")
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Procedure != "Unsafe" {
		t.Fatalf("default-item receiver findings = %+v, want only Unsafe", got)
	}
}

func TestVBA202Issue448PropagatesModuleFieldInitializationFromCallee(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private sharedSheet As Worksheet

Private Sub InitializeModuleState()
  Set sharedSheet = ThisWorkbook.Worksheets(1)
End Sub

Public Sub UseModuleState()
  InitializeModuleState
  Debug.Print sharedSheet.Name
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("module initializer should establish the field before use: %+v", got)
	}
}

func TestVBA202Issue448RefinesNotNothingGuard(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Guard(ByVal candidate As Worksheet)
  If Not candidate Is Nothing Then
    Debug.Print candidate.Name
  End If
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("Not-Is-Nothing guard should refine the true branch: %+v", got)
	}
}

func TestVBA202Issue448HonorsLocalShadowingOfModuleObject(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private sharedSheet As Worksheet

Public Sub Shadow()
  Dim sharedSheet As Worksheet
  Set Main.sharedSheet = ThisWorkbook.Worksheets(1)
  Debug.Print sharedSheet.Name
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Procedure != "Shadow" {
		t.Fatalf("shadowing local should remain nullable despite module initialization: %+v", got)
	}
}

func TestVBA202Issue448TracksExplicitNothingReset(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub ResetThenUse()
  Dim target As Worksheet
  Set target = ThisWorkbook.Worksheets(1)
  Set target = Nothing
  Debug.Print target.Name
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Line != 6 {
		t.Fatalf("explicit Nothing reset should invalidate later use: %+v", got)
	}
}
