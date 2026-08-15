package analyze

import (
	"encoding/json"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestVBA249DetectsLiteralAndProjectConstantRuntimeFailures(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Const ZeroValue As Long = 0
Public Const BadText As String = "not numeric"

Public Sub Run()
  Dim result As Double
  result = 10 / 0
  result = 10 / ZeroValue
  result = 10 \ 0
  result = 10 Mod ZeroValue
  result = BadText + 1
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA249")
	if len(got) != 5 {
		t.Fatalf("VBA249 findings = %+v, want deterministic division and numeric failures", got)
	}
	if got[0].Severity != "error" || !strings.Contains(got[0].Message, "guaranteed") {
		t.Fatalf("unexpected VBA249 finding = %+v", got[0])
	}
	if got[0].RuntimeError == nil || got[0].RuntimeError.Kind != "division_by_zero" {
		t.Fatalf("unexpected runtime error context = %+v", got[0].RuntimeError)
	}
	encoded, err := json.Marshal(got[0])
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(encoded), `"runtime_error":{"kind":"division_by_zero"}`) {
		t.Fatalf("runtime error context missing from JSON: %s", encoded)
	}
}

func TestVBA249LeavesUnknownAndNumericStringOperandsSilent(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim result As Double
  Dim denominator As Double
  result = 10 / denominator
  result = "123" + 1
  result = "1,2" + 1
  result = MissingText + 1
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 0 {
		t.Fatalf("unknown or numeric-string operands should remain silent: %+v", got)
	}
}

func TestVBA249CanBeDisabled(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Public Sub Run()
  Debug.Print 10 / 0
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectDeterministicRuntimeErrors = false
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 0 {
		t.Fatalf("disabled VBA249 produced findings: %+v", got)
	}
}

func TestVBA249PropagatesProcedureConstantsAcrossAllBranches(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal chooseFirst As Boolean)
  Dim result As Double
  Dim denominator As Double
  If chooseFirst Then
    denominator = 0
  Else
    denominator = 0
  End If
  result = 10 / denominator
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 1 {
		t.Fatalf("zero denominator on every branch should be reported once: %+v", got)
	}
}

func TestVBA249LeavesConflictingProcedureConstantsSilent(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal chooseFirst As Boolean)
  Dim result As Double
  Dim denominator As Double
  If chooseFirst Then
    denominator = 0
  Else
    denominator = 1
  End If
  result = 10 / denominator
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 0 {
		t.Fatalf("conflicting branch values should remain silent: %+v", got)
	}
}

func TestVBA249ResolvesProcedureConstAndConstantBranchReachability(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Const ZeroValue As Long = 0
  Dim denominator As Long
  If False Then
    denominator = 1
  Else
    denominator = ZeroValue
  End If
  Debug.Print 10 / denominator
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 1 {
		t.Fatalf("constant unreachable branch and local Const should prove zero divisor: %+v", got)
	}
}

func TestVBA249InvalidatesKnownValueAfterByRefMutation(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub Mutate(ByRef value As Long)
  value = 1
End Sub

Public Sub Run()
  Dim denominator As Long
  denominator = 0
  Mutate denominator
  Debug.Print 10 / denominator
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 0 {
		t.Fatalf("ByRef mutation must invalidate the zero fact: %+v", got)
	}
}

func TestVBA249BatchAndRealtimeResultsMatchAndRemainNonBlocking(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Option Explicit
Public Sub Run()
  Debug.Print 10 / 0
End Sub
`
	writeModule(t, dir, "Main.bas", source)
	cfg := config.Default()
	batch, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(dir, "src", "modules", "Main.bas")
	realtime, err := SourceRealtimeFindings(dir, path, cfg, []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	if got, want := findingsByCode(realtime, "VBA249"), findingsByCode(batch, "VBA249"); !reflect.DeepEqual(got, want) {
		t.Fatalf("batch/realtime VBA249 findings differ:\nbatch=%+v\nrealtime=%+v", want, got)
	}
	if blocking := BlockingFindings(findingsByCode(batch, "VBA249")); len(blocking) != 0 {
		t.Fatalf("runtime-error findings must not block preflight: %+v", blocking)
	}
}

func TestVBA249DetectsStrictNumericConversionFailures(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim value As Long
  value = CInt("not numeric")
  value = CInt(40000)
  value = CInt(32767.4)
  value = CInt("40000")
  value = CInt(12)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA249")
	if len(got) != 2 {
		t.Fatalf("expected only known conversion type/range failures, got %+v", got)
	}
	if got[0].RuntimeError == nil || got[1].RuntimeError == nil {
		t.Fatalf("conversion findings must carry runtime context: %+v", got)
	}
	for _, finding := range got {
		if finding.Line == 6 {
			t.Fatalf("banker's-rounded CInt(32767.4) must remain valid: %+v", got)
		}
	}
	kinds := []string{got[0].RuntimeError.Kind, got[1].RuntimeError.Kind}
	if !strings.Contains(strings.Join(kinds, ","), "conversion_type_mismatch") || !strings.Contains(strings.Join(kinds, ","), "conversion_overflow") {
		t.Fatalf("unexpected conversion kinds: %v", kinds)
	}
}

func TestVBA249DetectsNestedNumericConversionFailure(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim value As Long
  value = 1 + CInt(40000)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA249")
	if len(got) != 1 || got[0].RuntimeError == nil || got[0].RuntimeError.Kind != "conversion_overflow" {
		t.Fatalf("nested conversion overflow should be reported once: %+v", got)
	}
}

func TestVBA249DetectsKnownArrayBoundsAndKeepsUnknownShapesSilent(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal chooseFirst As Boolean)
  Dim fixed(0 To 1) As Long
  Dim values() As Long
  fixed(2) = 1
  If chooseFirst Then
    ReDim values(0 To 1)
  Else
    values = ExternalArray()
  End If
  values(2) = 1
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA249")
	if len(got) != 1 || got[0].RuntimeError == nil || got[0].RuntimeError.Kind != "array_subscript_out_of_bounds" {
		t.Fatalf("only the fixed-array out-of-bounds access should be deterministic: %+v", got)
	}
}
