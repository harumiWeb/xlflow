package analyze

import (
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestVBA257ReportsStandaloneFunctionCalls(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function LoadConfig() As Boolean
  LoadConfig = True
End Function

Public Function Lookup(ByVal key As String) As Long
  Lookup = 1
End Function

Public Sub Helper()
End Sub

Public Sub Run()
  LoadConfig
  Lookup "a"
  Call LoadConfig()
  LoadConfig()
  Helper
End Sub
`)
	cfg := config.Default()
	if findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run(); err != nil {
		t.Fatal(err)
	} else if got := findingsByCode(findings, "VBA257"); len(got) != 0 {
		t.Fatalf("VBA257 is opt-in; findings = %+v", got)
	}
	cfg.Analyze.DetectDiscardedFunctionReturn = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA257")
	if len(got) != 4 {
		t.Fatalf("VBA257 findings = %+v, want 4 discarded call sites", got)
	}
	for _, finding := range got {
		if finding.Severity != "warning" || finding.Procedure != "Run" || finding.File != "src/modules/Main.bas" {
			t.Fatalf("unexpected VBA257 finding: %+v", finding)
		}
	}
}

func TestVBA257SkipsConsumedResultsAndNonFunctionCallees(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function LoadConfig() As Boolean
  LoadConfig = True
End Function

Public Function MakeObject() As Object
  Set MakeObject = New Collection
End Function

Public Function Values(ByVal index As Long) As Long
  Values = index
End Function

Public Sub Consume(ByVal value As Variant)
End Sub

Public Sub Helper()
End Sub

Public Sub Run()
  Dim flag As Boolean
  Dim obj As Object
  Dim direct As Boolean
  flag = LoadConfig()
  Set obj = MakeObject()
  If LoadConfig() Then
    Helper
  End If
  Consume LoadConfig()
  Consume Values(1)
  Debug.Print LoadConfig()
  Debug.Print Values(2)
  direct = LoadConfig
  Helper
  MissingOperation
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectDiscardedFunctionReturn = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA257"); len(got) != 0 {
		t.Fatalf("consumed or unresolved calls must stay silent: %+v", got)
	}
}

func TestVBA257SkipsIndexedAssignmentAndRaiseEvent(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function Slots(ByVal size As Long) As Long()
  Dim result() As Long
  ReDim result(size)
  Slots = result
End Function

Public Sub Run()
  Dim values() As Long
  values = Slots(3)
  Slots(0) = 1
End Sub
`)
	classes := filepath.Join(dir, "src", "classes")
	if err := os.MkdirAll(classes, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(classes, "Events.cls"), []byte(`Option Explicit
Public Event Changed()

Public Function Compute() As Long
  Compute = 1
End Function

Public Sub Run()
  RaiseEvent Changed
  Me.Compute
End Sub
`), 0o644); err != nil {
		t.Fatal(err)
	}
	cfg := config.Default()
	cfg.Analyze.DetectDiscardedFunctionReturn = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA257")
	for _, finding := range got {
		if finding.Procedure == "Run" && finding.Module == "Main" {
			t.Fatalf("indexed assignment reported as discarded call: %+v", finding)
		}
	}
	found := false
	for _, finding := range got {
		if finding.Module == "Events" && finding.Procedure == "Run" {
			found = true
		}
	}
	if !found {
		t.Fatalf("Me.Compute member call to a resolved Function should report VBA257: %+v", got)
	}
}

func TestVBA257ReportsQualifiedCallAndSkipsWithMemberCall(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Mod.bas", `Option Explicit
Public Function LoadConfig() As Boolean
  LoadConfig = True
End Function
`)
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal holder As Object)
  Mod.LoadConfig
  With holder
    .LoadConfig
  End With
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectDiscardedFunctionReturn = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA257")
	if len(got) != 1 {
		t.Fatalf("qualified Mod.LoadConfig should report once; With member calls stay silent: %+v", got)
	}
	if got[0].Line != 3 {
		t.Fatalf("unexpected VBA257 location: %+v", got[0])
	}
}

func TestVBA257MatchesBatchAndRealtime(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function LoadConfig() As Boolean
  LoadConfig = True
End Function

Public Sub Run()
  LoadConfig
  Dim flag As Boolean
  flag = LoadConfig()
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectDiscardedFunctionReturn = true
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
	if got, want := findingsByCode(realtime, "VBA257"), findingsByCode(batch, "VBA257"); !reflect.DeepEqual(got, want) {
		t.Fatalf("batch/realtime VBA257 mismatch: batch=%+v realtime=%+v", want, got)
	}
}

func TestVBA258ReportsPrivateFunctionWhoseResultIsAlwaysDiscarded(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function LoadConfig() As Boolean
  LoadConfig = True
End Function

Public Function PublicLoad() As Boolean
  PublicLoad = True
End Function

Private Function Consumed() As Long
  Consumed = 1
End Function

Private Function NeverCalled() As Long
  NeverCalled = 1
End Function

Public Sub Run()
  LoadConfig
  PublicLoad
  Dim value As Long
  value = Consumed()
End Sub
`)
	cfg := config.Default()
	if findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run(); err != nil {
		t.Fatal(err)
	} else if got := findingsByCode(findings, "VBA258"); len(got) != 0 {
		t.Fatalf("VBA258 is opt-in; findings = %+v", got)
	}
	cfg.Analyze.DetectFunctionReturnAlwaysDiscarded = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA258")
	if len(got) != 1 {
		t.Fatalf("VBA258 findings = %+v, want only LoadConfig", got)
	}
	if got[0].Procedure != "LoadConfig" || got[0].Severity != "information" {
		t.Fatalf("unexpected VBA258 finding: %+v", got[0])
	}
}

func TestVBA258StaysSilentOnBareReferenceAndDynamicDispatch(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function LoadConfig() As Boolean
  LoadConfig = True
End Function

Private Function Referenced() As Boolean
  Referenced = True
End Function

Private Function Dynamic() As Boolean
  Dynamic = True
End Function

Public Sub Run()
  LoadConfig
  Referenced
  Dynamic
  Dim flag As Boolean
  flag = Referenced
  CallByName Me, "Dynamic", VbMethod
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectFunctionReturnAlwaysDiscarded = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA258")
	for _, finding := range got {
		if finding.Procedure == "Referenced" || finding.Procedure == "Dynamic" {
			t.Fatalf("non-call or dynamic reference must suppress VBA258: %+v", finding)
		}
	}
	if len(got) != 1 || got[0].Procedure != "LoadConfig" {
		t.Fatalf("VBA258 findings = %+v, want only LoadConfig", got)
	}
}

func TestVBA258StaysSilentOnUnknownDynamicTargetAndImplements(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function LoadConfig() As Boolean
  LoadConfig = True
End Function

Public Sub Run(ByVal macro As String)
  LoadConfig
  Application.Run macro
End Sub
`)
	classes := filepath.Join(dir, "src", "classes")
	if err := os.MkdirAll(classes, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(classes, "Impl.cls"), []byte(`Option Explicit
Implements IWorker

Private Function IWorker_DoWork() As Boolean
  IWorker_DoWork = True
End Function
`), 0o644); err != nil {
		t.Fatal(err)
	}
	cfg := config.Default()
	cfg.Analyze.DetectFunctionReturnAlwaysDiscarded = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA258"); len(got) != 0 {
		t.Fatalf("unknown dynamic target must suppress all VBA258 findings: %+v", got)
	}
}

func TestVBA258StaysSilentOnFunctionReferenceAsArgument(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "LibMemory.bas", `Option Explicit
Public Sub RedirectInstance(ByRef funcReturn As Variant, ByVal funcReturnPtr As Long, ByVal originalInstance As Object, ByVal newInstance As Object)
End Sub
`)
	classes := filepath.Join(dir, "src", "classes")
	if err := os.MkdirAll(classes, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(classes, "DemoClass.cls"), []byte(`Option Explicit
Private Function Init() As Boolean
  Init = True
End Function

Private Function OnlyCalled() As Boolean
  OnlyCalled = True
End Function

Public Sub Run()
  Init
  OnlyCalled
  RedirectInstance Init, VarPtr(Init), Me, Nothing
End Sub
`), 0o644); err != nil {
		t.Fatal(err)
	}
	cfg := config.Default()
	cfg.Analyze.DetectFunctionReturnAlwaysDiscarded = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA258")
	if len(got) != 1 || got[0].Procedure != "OnlyCalled" {
		t.Fatalf("an argument-position function reference inside another call must suppress VBA258: %+v", got)
	}
}

func TestVBA258ReportsWhenOnlyOwnReturnSlotIsReferenced(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "LibMemory.bas", `Option Explicit
Public Sub RedirectInstance(ByRef funcReturn As Variant, ByVal funcReturnPtr As Long, ByVal originalInstance As Object, ByVal newInstance As Object)
End Sub
`)
	classes := filepath.Join(dir, "src", "classes")
	if err := os.MkdirAll(classes, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(classes, "DemoClass.cls"), []byte(`Option Explicit
Private Function Init(ByVal c As DemoClass) As Boolean
  RedirectInstance Init, VarPtr(Init), Me, c
End Function

Public Sub Run()
  Init Nothing
End Sub
`), 0o644); err != nil {
		t.Fatal(err)
	}
	cfg := config.Default()
	cfg.Analyze.DetectFunctionReturnAlwaysDiscarded = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA258")
	if len(got) != 1 || got[0].Procedure != "Init" {
		t.Fatalf("own-body return-slot references are not call-site evidence; Init should report: %+v", got)
	}
}

func TestVBA257ReportsSingleLineIfAndColonSeparatedCalls(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function LoadConfig() As Boolean
  LoadConfig = True
End Function

Public Sub Helper()
End Sub

Public Sub Run(ByVal flag As Boolean)
  If flag Then LoadConfig
  Helper : LoadConfig
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectDiscardedFunctionReturn = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA257")
	if len(got) != 2 {
		t.Fatalf("single-line If and colon-separated standalone calls discard the result: %+v", got)
	}
	if got[0].Line != 10 || got[1].Line != 11 {
		t.Fatalf("unexpected VBA257 locations: %+v", got)
	}
}
func TestVBA257SkipsGrammarMisSplitCallStatement(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "WebHelpers.bas", `Option Explicit
Public Function MethodToName(ByVal method As Long) As String
  MethodToName = "GET"
End Function
`)
	classes := filepath.Join(dir, "src", "classes")
	if err := os.MkdirAll(classes, 0o755); err != nil {
		t.Fatal(err)
	}
	// The grammar can split `web_Http.Open arg, arg, arg` into an expression
	// statement plus a bogus call statement when the member name is a
	// reserved keyword. The bogus call makes MethodToName look discarded
	// even though web_Http.Open consumes its result.
	if err := os.WriteFile(filepath.Join(classes, "WebClient.cls"), []byte("Option Explicit\n"+
		"Public Function GetFullUrl(Request As Object) As String\n"+
		"  GetFullUrl = \"url\"\n"+
		"End Function\n"+
		"Public Function PrepareHttpRequest(Request As Object, Optional Async As Boolean = True) As Object\n"+
		"    Dim web_Http As Object\n"+
		"    On Error GoTo web_ErrorHandling\n"+
		"    Set web_Http = CreateObject(\"WinHttp.WinHttpRequest.5.1\")\n"+
		"    web_BeforeExecute Request\n"+
		"    ' Open http request\n"+
		"    web_Http.Open WebHelpers.MethodToName(Request.Method), Me.GetFullUrl(Request), Async\n"+
		"    web_Http.SetTimeouts 1, 2, 3, 4\n"+
		"    Exit Function\n"+
		"web_ErrorHandling:\n"+
		"End Function\n"+
		"Private Sub web_BeforeExecute(Request As Object)\n"+
		"End Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	cfg := config.Default()
	cfg.Analyze.DetectDiscardedFunctionReturn = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA257"); len(got) != 0 {
		t.Fatalf("mis-split argument calls must not report VBA257: %+v", got)
	}
}

func TestVBA258StaysSilentOnSameNameArgument(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function Foo(ByRef slot As Variant) As Boolean
  Foo = True
End Function

Public Sub Run()
  Foo Foo
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectFunctionReturnAlwaysDiscarded = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	// The argument Foo is a value reference to the function, not callee
	// syntax, so the all-discard claim must stay silent.
	if got := findingsByCode(findings, "VBA258"); len(got) != 0 {
		t.Fatalf("same-named argument reference must suppress VBA258: %+v", got)
	}
}

func TestVBA258StaysSilentOnImplicitApplicationDispatch(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function LoadConfig() As Boolean
  LoadConfig = True
End Function

Private Function OtherTask() As Boolean
  OtherTask = True
End Function

Private Function ThirdTask() As Boolean
  ThirdTask = True
End Function

Public Sub Driver()
  LoadConfig
  OtherTask
  ThirdTask
  Run "LoadConfig"
  Dim result As Variant
  result = Run("OtherTask")
  With Application
    .Run "ThirdTask"
  End With
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectFunctionReturnAlwaysDiscarded = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	// Receiverless Run and With-block .Run are implicit Application dispatch:
	// their string arguments name procedures the static call set cannot see.
	if got := findingsByCode(findings, "VBA258"); len(got) != 0 {
		t.Fatalf("implicit Application dispatch targets must suppress VBA258: %+v", got)
	}
}

func TestVBA258StaysSilentOnInterfaceImplementation(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	classes := filepath.Join(dir, "src", "classes")
	if err := os.MkdirAll(classes, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(classes, "Impl.cls"), []byte(`Option Explicit
Implements IWorker

Private Function IWorker_DoWork() As Boolean
  IWorker_DoWork = True
End Function

Private Function Parse_Name() As Boolean
  Parse_Name = True
End Function

Public Sub Drive()
  IWorker_DoWork
  Parse_Name
End Sub
`), 0o644); err != nil {
		t.Fatal(err)
	}
	cfg := config.Default()
	cfg.Analyze.DetectFunctionReturnAlwaysDiscarded = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA258")
	for _, finding := range got {
		if finding.Procedure == "IWorker_DoWork" {
			t.Fatalf("interface implementations are invoked through the interface: %+v", finding)
		}
	}
	// Parse_Name is an unrelated underscored helper, not an interface
	// implementation, so its all-discard call set still reports.
	if len(got) != 1 || got[0].Procedure != "Parse_Name" {
		t.Fatalf("unrelated underscored helper keeps VBA258 coverage: %+v", got)
	}
}

func TestVBA258StaysSilentOnUnderscoredInterfaceImplementation(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	classes := filepath.Join(dir, "src", "classes")
	if err := os.MkdirAll(classes, 0o755); err != nil {
		t.Fatal(err)
	}
	// Interface names may contain underscores: Implements I_Worker names the
	// DoWork implementation I_Worker_DoWork. Matching only the segment before
	// the first underscore would miss the prefix and keep the implementation
	// eligible even though interface-typed callers can consume its result.
	if err := os.WriteFile(filepath.Join(classes, "Impl.cls"), []byte(`Option Explicit
Implements I_Worker

Private Function I_Worker_DoWork() As Boolean
  I_Worker_DoWork = True
End Function

Public Sub Drive()
  I_Worker_DoWork
End Sub
`), 0o644); err != nil {
		t.Fatal(err)
	}
	cfg := config.Default()
	cfg.Analyze.DetectFunctionReturnAlwaysDiscarded = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, finding := range findingsByCode(findings, "VBA258") {
		if finding.Procedure == "I_Worker_DoWork" {
			t.Fatalf("interface implementations are invoked through the interface: %+v", finding)
		}
	}
}

func TestVBA258StaysSilentOnParseError(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function LoadConfig() As Boolean
  LoadConfig = True
End Function

Public Sub Run()
  LoadConfig
End Sub
`)
	// A second Dim keyword after a declaration comma is a recovery-accepted
	// parse error: the file still reaches analysis with HasError set.
	writeModule(t, dir, "Broken.bas", `Option Explicit
Public Sub Helper()
  Dim x As Long, Dim y As Long
  y = x
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectFunctionReturnAlwaysDiscarded = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	// A parse error anywhere can hide a consuming call or a dynamic-dispatch
	// reference, so no all-discard claim survives.
	if got := findingsByCode(findings, "VBA258"); len(got) != 0 {
		t.Fatalf("parse errors must fail open: %+v", got)
	}
}

func TestVBA258StaysSilentWhenPathFilterHidesModules(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Friend Function LoadConfig() As Boolean
  LoadConfig = True
End Function

Public Sub Run()
  LoadConfig
End Sub
`)
	writeModule(t, dir, "Hidden.bas", `Option Explicit
Public Sub Consume()
  Dim value As Boolean
  value = LoadConfig()
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectFunctionReturnAlwaysDiscarded = true
	findings, err := (Analyzer{
		RootDir: dir,
		Config:  cfg,
		PathFilter: func(path string) bool {
			return strings.EqualFold(filepath.Base(path), "Main.bas")
		},
	}).Run()
	if err != nil {
		t.Fatal(err)
	}
	// Hidden.bas consumes the result, but PathFilter removes it from the
	// project view; filtered-out modules can also hide dynamic references,
	// so the all-discard claim must stay silent.
	if got := findingsByCode(findings, "VBA258"); len(got) != 0 {
		t.Fatalf("filtered project view must suppress VBA258: %+v", got)
	}
}
