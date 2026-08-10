package analyze

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestDictionaryCollectionSafetyRules(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub DictionaryRisks(ByVal key As String)
  Dim dict As Object
  Dim i As Long
  Set dict = CreateObject("Scripting.Dictionary")
  dict.Add LCase$(key), 1
  dict.CompareMode = TextCompare
  For i = 1 To 3
    Debug.Print dict.Keys()(0)
  Next i
  Debug.Print dict(key)
End Sub

Public Sub CollectionRisks()
  Dim values As Collection
  Dim item As Variant
  Set values = New Collection
  values.Add "value", "key"
  For Each item In values
    values.Remove "key"
  Next item
  Debug.Print values(0)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	for code, line := range map[string]int{
		"VBA230": 7,
		"VBA231": 9,
		"VBA232": 11,
		"VBA233": 7,
		"VBA234": 20,
		"VBA235": 22,
	} {
		assertFinding(t, findings, code, line)
	}
}

func TestVBA207TracksDefiniteAndUnknownKeyState(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal external As Scripting.Dictionary, ByVal key As String)
  Dim localDict As Object
  Set localDict = CreateObject("Scripting.Dictionary")
  Debug.Print localDict(key)
  If localDict.Exists(key) Then
    Debug.Print localDict(key)
  End If
  Debug.Print external(key)
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectDictionaryCollectionGuard = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA207")
	if len(got) != 2 {
		t.Fatalf("VBA207 findings = %+v, want definite and unknown only", got)
	}
	if got[0].Line != 5 || got[0].Severity != "warning" {
		t.Fatalf("definitely absent VBA207 = %+v", got[0])
	}
	if got[1].Line != 9 || got[1].Severity != "information" {
		t.Fatalf("unknown VBA207 = %+v", got[1])
	}
}

func TestVBA207KeepsBranchJoinedContentUnknown(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal clearValues As Boolean)
  Dim dict As Scripting.Dictionary
  Set dict = New Scripting.Dictionary
  dict.Add "key", 1
  If clearValues Then
    dict.RemoveAll
  End If
  Debug.Print dict("key")
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectDictionaryCollectionGuard = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA207")
	if len(got) != 1 || got[0].Line != 9 || got[0].Severity != "information" {
		t.Fatalf("branch-joined VBA207 = %+v, want one information finding", got)
	}
}

func TestDictionaryCollectionOverlayPreservesFlowContentFacts(t *testing.T) {
	t.Parallel()
	flow := &dcFlowState{
		Bindings: map[string]string{"dict": "object"},
		Objects: map[string]dcObjectState{
			"object": {Kind: dcDictionary, Empty: 0, Keys: map[string]int{"key|": 0}, Normalizations: map[string]string{"key": ""}},
		},
		Scalars: map[string]dcKeyExpr{},
	}
	linear := &dcFlowState{
		Bindings: map[string]string{"dict": "object"},
		Objects: map[string]dcObjectState{
			"object": {Kind: dcDictionary, Empty: 1, Keys: map[string]int{}, Normalizations: map[string]string{}},
		},
		Scalars: map[string]dcKeyExpr{},
	}

	got := dcOverlayState(flow, linear).Objects["object"]
	keyState, keyExists := got.Keys["key|"]
	if got.Empty != 0 || !keyExists || keyState != 0 {
		t.Fatalf("overlay content facts = %+v, want CFG facts", got)
	}
}

func TestVBA207DoesNotInferFactsFromCompoundExistsCondition(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal dict As Scripting.Dictionary, ByVal key As String, ByVal useDefault As Boolean)
  If dict.Exists(key) Or useDefault Then
    Debug.Print "selected"
  Else
    Debug.Print dict(key)
  End If
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectDictionaryCollectionGuard = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA207")
	if len(got) != 1 || got[0].Line != 6 || got[0].Severity != "information" {
		t.Fatalf("compound Exists VBA207 = %+v, want one information finding", got)
	}
}

func TestDictionaryCollectionSafetyAcceptsSafePatterns(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal key As String)
  Dim dict As Object
  Dim values As Collection
  Dim keys As Variant
  Dim i As Long
  Set dict = CreateObject("Scripting.Dictionary")
  dict.CompareMode = vbTextCompare
  dict.Add LCase$(key), 1
  If dict.Exists(LCase$(key)) Then Debug.Print dict(LCase$(key))
  keys = dict.Keys
  For i = LBound(keys) To UBound(keys)
    Debug.Print keys(i)
  Next i
  Set values = New Collection
  values.Add "value", "key"
  For i = 1 To values.Count
    Debug.Print values(i)
  Next i
  On Error Resume Next
  Debug.Print values("optional")
  On Error GoTo 0
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectDictionaryCollectionGuard = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, code := range []string{"VBA207", "VBA230", "VBA231", "VBA232", "VBA233", "VBA234", "VBA235"} {
		if got := findingsByCode(findings, code); len(got) != 0 {
			t.Fatalf("safe patterns reported %s: %+v", code, got)
		}
	}
}

func TestDictionaryCollectionSafetyTracksAliasesAndHelpers(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function NewDictionary() As Object
  Set NewDictionary = CreateObject("Scripting.Dictionary")
End Function

Private Sub RemoveValue(ByVal values As Collection, ByVal key As String)
  values.Remove key
End Sub

Public Sub Run()
  Dim original As Object
  Dim aliasDict As Object
  Dim item As Variant
  Dim values As Collection
  Dim i As Long
  Set original = NewDictionary()
  Set aliasDict = original
  For Each item In aliasDict
    Debug.Print item.Name
  Next item
  Set values = New Collection
  values.Add "value", "key"
  For Each item In values
    RemoveValue values, "key"
  Next item
  For i = 0 To values.Count - 1
    Debug.Print values(i)
  Next i
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectDictionaryIterationValueUsage = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	assertFinding(t, findings, "VBA213", 19)
	assertFinding(t, findings, "VBA234", 24)
	assertFinding(t, findings, "VBA235", 27)
}

func TestVBA231DoesNotPropagateUntypedKeysHelper(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub WriteUnknown(ByVal value As Object)
  Debug.Print value.Keys
End Sub

Public Sub Run(ByVal values As Collection)
  Dim item As Variant
  For Each item In values
    WriteUnknown item
  Next item
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA231"); len(got) != 0 {
		t.Fatalf("untyped Keys helper must remain unknown: %+v", got)
	}
}

func TestQualifiedMemberCallDoesNotResolveProjectHelper(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub Remove(ByVal values As Collection)
  values.Remove 1
End Sub

Public Sub Run(ByVal dict As Scripting.Dictionary, ByVal values As Collection)
  Dim item As Variant
  For Each item In values
    dict.Remove values
  Next item
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA234"); len(got) != 0 {
		t.Fatalf("qualified member call resolved project helper: %+v", got)
	}
}

func TestVBA234AllowsMutationBeforeUnconditionalLoopExit(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub ReplaceAndExit(ByVal values As Collection)
  Dim item As Variant
  For Each item In values
    If item = "target" Then
      values.Remove 1
      values.Add "replacement"
      Exit Sub
    End If
  Next item
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA234"); len(got) != 0 {
		t.Fatalf("mutation followed by unconditional procedure exit is safe for the enumerator: %+v", got)
	}
}

func TestVBA235TracksExplicitZeroBasedArrayIndex(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub CopyValues(ByVal values As Collection)
  Dim buffer(0 To 3) As Variant
  Dim i As Long
  For i = LBound(buffer) To UBound(buffer)
    buffer(i) = values(i)
  Next i
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	assertFinding(t, findings, "VBA235", 6)
}

func TestDictionaryCollectionSafetyIgnoresCommentsStringsAndInlineSuppression(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim dict As Scripting.Dictionary
  Dim values As Collection
  Dim i As Long
  Set dict = New Scripting.Dictionary
  dict.Add "key", 1
  dict.CompareMode = vbTextCompare ' xlflow:disable-line VBA230
  Set values = New Collection
  For i = 1 To 2
    Debug.Print "dict.Keys and values(0)"
    ' values.Remove 1
  Next i
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, code := range []string{"VBA230", "VBA231", "VBA234", "VBA235"} {
		if got := findingsByCode(findings, code); len(got) != 0 {
			t.Fatalf("comments, strings, or suppression reported %s: %+v", code, got)
		}
	}
}
