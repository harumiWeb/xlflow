package analyze

import (
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func runOptionBaseAnalysis(t *testing.T, modules map[string]string, configure func(*config.Config)) []Finding {
	t.Helper()
	dir := t.TempDir()
	for name, source := range modules {
		writeModule(t, dir, name, source)
	}
	cfg := config.Default()
	if configure != nil {
		configure(&cfg)
	}
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	return findings
}

func enableOptionBaseRules(cfg *config.Config) {
	cfg.Analyze.DetectOptionBaseArrayInconsistency = true
	cfg.Analyze.DetectOptionBaseParamArrayInconsistency = true
}

// ---------- VBA270: Array call ignores Option Base 1 ----------

func TestOptionBaseArrayReportsQualifiedVBAArray(t *testing.T) {
	findings := runOptionBaseAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Option Base 1
Public Sub Run()
    Dim values As Variant
    values = VBA.Array("a", "b", "c")
End Sub
`}, enableOptionBaseRules)
	got := findingsByCode(findings, "VBA270")
	if len(got) != 1 {
		t.Fatalf("VBA270 findings = %+v, want one", got)
	}
	if !strings.Contains(got[0].Message, "VBA.Array") {
		t.Fatalf("VBA270 message = %q, want callee text", got[0].Message)
	}
}

func TestOptionBaseArrayReportsUnqualifiedIntrinsic(t *testing.T) {
	findings := runOptionBaseAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Option Base 1
Public Sub Run()
    Dim values As Variant
    values = Array("a", "b", "c")
End Sub
`}, enableOptionBaseRules)
	got := findingsByCode(findings, "VBA270")
	if len(got) != 1 {
		t.Fatalf("VBA270 findings = %+v, want one", got)
	}
}

func TestOptionBaseArrayReportsNestedCall(t *testing.T) {
	findings := runOptionBaseAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Option Base 1
Public Sub Run()
    Debug.Print LBound(Array(1, 2))
End Sub
`}, enableOptionBaseRules)
	got := findingsByCode(findings, "VBA270")
	if len(got) != 1 {
		t.Fatalf("VBA270 findings = %+v, want one for nested Array call", got)
	}
}

func TestOptionBaseArraySkipsUserDefinedArrayFunction(t *testing.T) {
	findings := runOptionBaseAnalysis(t, map[string]string{
		"Main.bas": `Option Explicit
Option Base 1
Public Sub Run()
    Dim values As Variant
    values = Array("a", "b", "c")
End Sub
`,
		"Helpers.bas": `Option Explicit
Public Function Array(ParamArray items() As Variant) As Variant
    Array = items
End Function
`}, enableOptionBaseRules)
	if got := findingsByCode(findings, "VBA270"); len(got) != 0 {
		t.Fatalf("VBA270 findings = %+v, want none for user-defined Array", got)
	}
}

func TestOptionBaseArraySkipsLocalShadow(t *testing.T) {
	findings := runOptionBaseAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Option Base 1
Public Sub Run()
    Dim Array(0 To 2) As String
    Dim first As String
    Array(0) = "a"
    first = Array(0)
End Sub
`}, enableOptionBaseRules)
	if got := findingsByCode(findings, "VBA270"); len(got) != 0 {
		t.Fatalf("VBA270 findings = %+v, want none for lexically shadowed Array", got)
	}
}

func TestOptionBaseArraySkipsModuleShadow(t *testing.T) {
	findings := runOptionBaseAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Option Base 1
Private Array(0 To 2) As String
Public Sub Run()
    Dim first As String
    first = Array(0)
End Sub
`}, enableOptionBaseRules)
	if got := findingsByCode(findings, "VBA270"); len(got) != 0 {
		t.Fatalf("VBA270 findings = %+v, want none for module-level shadowed Array", got)
	}
}

func TestOptionBaseArraySkipsNonVBAReceiver(t *testing.T) {
	findings := runOptionBaseAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Option Base 1
Public Sub Run()
    Dim values As Variant
    values = Helpers.Array("a")
End Sub
`}, enableOptionBaseRules)
	if got := findingsByCode(findings, "VBA270"); len(got) != 0 {
		t.Fatalf("VBA270 findings = %+v, want none for a non-VBA qualified receiver", got)
	}
}

func TestOptionBaseArraySkipsSameModuleArrayFunction(t *testing.T) {
	findings := runOptionBaseAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Option Base 1
Public Function Array(ByVal seed As Long) As Variant
    Array = VBA.Array(seed)
End Function
Public Sub Run()
    Dim values As Variant
    values = Array(1)
End Sub
`}, enableOptionBaseRules)
	got := findingsByCode(findings, "VBA270")
	if len(got) != 1 {
		t.Fatalf("VBA270 findings = %+v, want only the explicit VBA.Array call, not the user-defined Array", got)
	}
	if !strings.Contains(got[0].Message, "VBA.Array") {
		t.Fatalf("VBA270 message = %q, want the qualified intrinsic call", got[0].Message)
	}
}

func TestOptionBaseArraySkipsIndexedAssignmentOnly(t *testing.T) {
	findings := runOptionBaseAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Option Base 1
Public Sub Run()
    Array(0) = "a"
End Sub
`}, enableOptionBaseRules)
	if got := findingsByCode(findings, "VBA270"); len(got) != 0 {
		t.Fatalf("VBA270 findings = %+v, want none for an indexed assignment target", got)
	}
}

func TestOptionBaseArraySkipsParameterShadow(t *testing.T) {
	findings := runOptionBaseAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Option Base 1
Public Sub Run(Array() As Variant)
    Dim first As Variant
    first = Array(0)
End Sub
`}, enableOptionBaseRules)
	if got := findingsByCode(findings, "VBA270"); len(got) != 0 {
		t.Fatalf("VBA270 findings = %+v, want none for a parameter shadowing Array", got)
	}
}

func TestOptionBaseSkipsConditionalBranchProcedure(t *testing.T) {
	findings := runOptionBaseAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Option Base 1
#If VBA7 Then
Public Sub Run(ParamArray values())
    Dim copy As Variant
    copy = Array("x")
End Sub
#End If
`}, enableOptionBaseRules)
	if got := findingsByCode(findings, "VBA270"); len(got) != 0 {
		t.Fatalf("VBA270 findings = %+v, want none inside conditional-compilation branches", got)
	}
	if got := findingsByCode(findings, "VBA271"); len(got) != 0 {
		t.Fatalf("VBA271 findings = %+v, want none inside conditional-compilation branches", got)
	}
}

func TestOptionBaseArraySilentWithoutOptionBase1(t *testing.T) {
	findings := runOptionBaseAnalysis(t, map[string]string{
		"Base0.bas": `Option Explicit
Option Base 0
Public Sub Run0()
    Dim values As Variant
    values = Array("a")
End Sub
`,
		"NoBase.bas": `Option Explicit
Public Sub Run1()
    Dim values As Variant
    values = VBA.Array("a")
End Sub
`}, enableOptionBaseRules)
	if got := findingsByCode(findings, "VBA270"); len(got) != 0 {
		t.Fatalf("VBA270 findings = %+v, want none under Option Base 0 / absent", got)
	}
}

func TestOptionBaseArraySkipsExplicitlyBoundedDeclarations(t *testing.T) {
	findings := runOptionBaseAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Option Base 1
Public Sub Run()
    Dim fixed(1 To 5) As String
    Dim sized(5) As String
    ReDim dynamic(1 To 3) As String
End Sub
`}, enableOptionBaseRules)
	if got := findingsByCode(findings, "VBA270"); len(got) != 0 {
		t.Fatalf("VBA270 findings = %+v, want none for bounded array declarations", got)
	}
}

func TestOptionBaseArrayDisabledByDefault(t *testing.T) {
	findings := runOptionBaseAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Option Base 1
Public Sub Run()
    Dim values As Variant
    values = Array("a")
End Sub
`}, nil)
	if got := findingsByCode(findings, "VBA270"); len(got) != 0 {
		t.Fatalf("VBA270 findings = %+v, want none when the rule is not enabled", got)
	}
}

// ---------- VBA271: ParamArray ignores Option Base 1 ----------

func TestOptionBaseParamArrayReportsParameter(t *testing.T) {
	findings := runOptionBaseAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Option Base 1
Public Sub Example(ParamArray values())
    Debug.Print LBound(values)
End Sub
`}, enableOptionBaseRules)
	got := findingsByCode(findings, "VBA271")
	if len(got) != 1 {
		t.Fatalf("VBA271 findings = %+v, want one", got)
	}
	if !strings.Contains(got[0].Message, "values") {
		t.Fatalf("VBA271 message = %q, want parameter name", got[0].Message)
	}
	// The diagnostic anchors on the `values` identifier inside
	// `Public Sub Example(ParamArray values())`, not the whole
	// `ParamArray values()` declaration.
	if got[0].Line != 3 || got[0].Column != 31 || got[0].EndLine != 3 || got[0].EndColumn != 37 {
		t.Fatalf("VBA271 range = line:%d col:%d-%d:%d, want 3:31-3:37 (identifier only)",
			got[0].Line, got[0].Column, got[0].EndLine, got[0].EndColumn)
	}
}

func TestOptionBaseParamArraySilentWithoutOptionBase1(t *testing.T) {
	findings := runOptionBaseAnalysis(t, map[string]string{
		"Base0.bas": `Option Explicit
Option Base 0
Public Sub Example0(ParamArray values())
End Sub
`,
		"NoBase.bas": `Option Explicit
Public Sub Example1(ParamArray values())
End Sub
`}, enableOptionBaseRules)
	if got := findingsByCode(findings, "VBA271"); len(got) != 0 {
		t.Fatalf("VBA271 findings = %+v, want none under Option Base 0 / absent", got)
	}
}

func TestOptionBaseParamArraySkipsOrdinaryParameters(t *testing.T) {
	findings := runOptionBaseAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Option Base 1
Public Sub Example(ByVal values() As String, Optional ByVal extra As Variant)
End Sub
`}, enableOptionBaseRules)
	if got := findingsByCode(findings, "VBA271"); len(got) != 0 {
		t.Fatalf("VBA271 findings = %+v, want none for ordinary array parameters", got)
	}
}

func TestOptionBaseMatchesBatchAndRealtime(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Option Base 1
Public Sub Run(ParamArray ignored())
    Dim values As Variant
    values = Array("a", "b")
End Sub
`)
	cfg := config.Default()
	enableOptionBaseRules(&cfg)
	path := filepath.Join(dir, "src", "modules", "Main.bas")
	source, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	batch, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	realtime, err := SourceRealtimeFindings(dir, path, cfg, source)
	if err != nil {
		t.Fatal(err)
	}
	for _, code := range []string{"VBA270", "VBA271"} {
		if got, want := findingsByCode(realtime, code), findingsByCode(batch, code); !reflect.DeepEqual(got, want) {
			t.Fatalf("batch/realtime %s mismatch: batch=%+v realtime=%+v", code, want, got)
		}
	}
}

func TestOptionBaseMixedConstructsReportIndependently(t *testing.T) {
	findings := runOptionBaseAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Option Base 1
Public Sub Example(ParamArray values())
    Dim copy As Variant
    copy = Array("x")
End Sub
`}, enableOptionBaseRules)
	if got := findingsByCode(findings, "VBA270"); len(got) != 1 {
		t.Fatalf("VBA270 findings = %+v, want one", got)
	}
	if got := findingsByCode(findings, "VBA271"); len(got) != 1 {
		t.Fatalf("VBA271 findings = %+v, want one", got)
	}
}
