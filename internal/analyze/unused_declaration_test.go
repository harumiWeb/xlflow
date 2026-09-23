package analyze

import (
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/typedb"
	"github.com/harumiWeb/xlflow/internal/vba/sourceproject"
)

func enableUnusedDeclarationRules(cfg config.Config) config.Config {
	cfg.Analyze.DetectUnusedParameters = true
	cfg.Analyze.DetectUnusedPrivateConstants = true
	cfg.Analyze.DetectUnusedUDTMembers = true
	cfg.Analyze.DetectNeverAssignedVariables = true
	cfg.Analyze.DetectUnassignedVariableUsage = true
	return cfg
}

func runUnusedDeclarationAnalysis(t *testing.T, modules map[string]string) []Finding {
	t.Helper()
	dir := t.TempDir()
	for name, source := range modules {
		writeModule(t, dir, name, source)
	}
	findings, err := (Analyzer{RootDir: dir, Config: enableUnusedDeclarationRules(config.Default())}).Run()
	if err != nil {
		t.Fatal(err)
	}
	return findings
}

// ---------- VBA260: unused procedure parameter ----------

func TestUnusedParameterReportsUnreadParameter(t *testing.T) {
	findings := runUnusedDeclarationAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Private Sub Helper(ByVal used As Long, ByVal extra As Long)
  Debug.Print used
End Sub
`})
	got := findingsByCode(findings, "VBA260")
	if len(got) != 1 || !strings.Contains(got[0].Message, "extra") {
		t.Fatalf("VBA260 findings = %+v, want one finding naming extra", got)
	}
}

func TestUnusedParameterSkipsUsedAndIgnoredNames(t *testing.T) {
	findings := runUnusedDeclarationAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Private Sub Helper(ByVal used As Long, ByVal unusedParam As Long)
  Debug.Print used
End Sub
`})
	if got := findingsByCode(findings, "VBA260"); len(got) != 0 {
		t.Fatalf("VBA260 findings = %+v, want none", got)
	}
}

func TestUnusedParameterSkipsPublicFriendAndDynamicEntryPoints(t *testing.T) {
	findings := runUnusedDeclarationAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Public Sub PublicHelper(ByVal arg As Long)
End Sub
Friend Sub FriendHelper(ByVal arg As Long)
End Sub
Private Sub DynamicHelper(ByVal arg As Long)
End Sub
Public Sub Runner()
  Application.Run "Main.DynamicHelper", 1
End Sub
`})
	if got := findingsByCode(findings, "VBA260"); len(got) != 0 {
		t.Fatalf("VBA260 findings = %+v, want none", got)
	}
}

func TestUnusedParameterSkipsEventHandlers(t *testing.T) {
	t.Setenv(typedb.EnvDir, t.TempDir())
	project := sourceproject.SourceProject{Files: []sourceproject.SourceFile{{
		Path: "src/workbook/Sheet1.cls",
		Source: []byte("Option Explicit\n" +
			"Private Sub Worksheet_Change(ByVal Target As Range)\n" +
			"End Sub\n"),
		ModuleKind: sourceproject.ModuleKindDocument,
	}}}
	result, err := (Analyzer{Config: enableUnusedDeclarationRules(config.Default())}).AnalyzeProject(t.Context(), project)
	if err != nil {
		t.Fatalf("AnalyzeProject: %v", err)
	}
	if got := findingsByCode(result.Findings, "VBA260"); len(got) != 0 {
		t.Fatalf("VBA260 findings = %+v, want none for a document event handler", got)
	}
}

func TestUnusedParameterSkipsWithEventsCallback(t *testing.T) {
	findings := runUnusedDeclarationAnalysis(t, map[string]string{"Handler.cls": `Option Explicit
Private WithEvents App As Application
Private Sub App_WindowBeforeDoubleClick(ByVal Sel As Selection, ByRef Cancel As Boolean)
  Cancel = True
End Sub
`})
	if got := findingsByCode(findings, "VBA260"); len(got) != 0 {
		t.Fatalf("VBA260 findings = %+v, want none for a WithEvents callback", got)
	}
}

func TestUnusedParameterSkipsDocumentObjectEventShape(t *testing.T) {
	t.Setenv(typedb.EnvDir, t.TempDir())
	project := sourceproject.SourceProject{Files: []sourceproject.SourceFile{{
		Path: "src/workbook/Report_Weekly.cls",
		Source: []byte("Option Explicit\n" +
			"Private Sub secDetail_Format(ByVal Cancel As Long, ByVal FormatCount As Long)\n" +
			"End Sub\n" +
			"Private Sub Helper()\n" +
			"End Sub\n"),
		ModuleKind: sourceproject.ModuleKindDocument,
	}}}
	result, err := (Analyzer{Config: enableUnusedDeclarationRules(config.Default())}).AnalyzeProject(t.Context(), project)
	if err != nil {
		t.Fatalf("AnalyzeProject: %v", err)
	}
	if got := findingsByCode(result.Findings, "VBA260"); len(got) != 0 {
		t.Fatalf("VBA260 findings = %+v, want none for a document object-event shape", got)
	}
}

func TestUnusedParameterSkipsImplementsMembers(t *testing.T) {
	findings := runUnusedDeclarationAnalysis(t, map[string]string{
		"IFace.cls": `Option Explicit
Public Sub DoWork(ByVal arg As Long)
End Sub
`,
		"Impl.cls": `Option Explicit
Implements IFace
Private Sub IFace_DoWork(ByVal arg As Long)
End Sub
`,
	})
	if got := findingsByCode(findings, "VBA260"); len(got) != 0 {
		t.Fatalf("VBA260 findings = %+v, want none", got)
	}
}

func TestUnusedParameterMatchesCaseInsensitively(t *testing.T) {
	findings := runUnusedDeclarationAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Private Sub Helper(ByVal Value As Long)
  Debug.Print VALUE
End Sub
`})
	if got := findingsByCode(findings, "VBA260"); len(got) != 0 {
		t.Fatalf("VBA260 findings = %+v, want none for case-insensitive use", got)
	}
}

// ---------- VBA261: unused private constant ----------

func TestUnusedPrivateConstReportsUnreadConstant(t *testing.T) {
	findings := runUnusedDeclarationAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Private Const UsedLimit As Long = 10
Private Const DeadLimit As Long = 20
Public Sub Run()
  Debug.Print UsedLimit
End Sub
`})
	got := findingsByCode(findings, "VBA261")
	if len(got) != 1 || !strings.Contains(got[0].Message, "DeadLimit") {
		t.Fatalf("VBA261 findings = %+v, want one finding naming DeadLimit", got)
	}
}

func TestUnusedPrivateConstCountsConditionalCompilationAndDeclarationUse(t *testing.T) {
	findings := runUnusedDeclarationAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Private Const FlagEnabled As Long = 1
Private Const SeedValue As Long = 7
Private Const DerivedLimit As Long = SeedValue + 1
#If FlagEnabled Then
Public Sub Run()
  Debug.Print DerivedLimit
End Sub
#End If
`})
	if got := findingsByCode(findings, "VBA261"); len(got) != 0 {
		t.Fatalf("VBA261 findings = %+v, want none", got)
	}
}

func TestUnusedPrivateConstSkipsEnumMembers(t *testing.T) {
	findings := runUnusedDeclarationAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Private Enum Mode
  Alpha
  Beta
End Enum
Public Sub Run()
  Debug.Print "done"
End Sub
`})
	if got := findingsByCode(findings, "VBA261"); len(got) != 0 {
		t.Fatalf("VBA261 findings = %+v, want none for enum members", got)
	}
}

// ---------- VBA262: unused UDT member ----------

func TestUnusedUDTMemberReportsUnreadMember(t *testing.T) {
	findings := runUnusedDeclarationAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Private Type TPoint
  X As Long
  Y As Long
End Type
Private Sub Helper()
  Dim p As TPoint
  p.X = 1
End Sub
`})
	got := findingsByCode(findings, "VBA262")
	if len(got) != 1 || !strings.Contains(got[0].Message, "Y") {
		t.Fatalf("VBA262 findings = %+v, want one finding naming Y", got)
	}
}

func TestUnusedUDTMemberCountsWithBlockAccess(t *testing.T) {
	findings := runUnusedDeclarationAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Private Type TPoint
  X As Long
  Y As Long
End Type
Private Sub Helper()
  Dim p As TPoint
  With p
    .X = 1
    .Y = 2
  End With
End Sub
`})
	if got := findingsByCode(findings, "VBA262"); len(got) != 0 {
		t.Fatalf("VBA262 findings = %+v, want none", got)
	}
}

func TestUnusedUDTMemberFailsOpenOnLateBoundReceiver(t *testing.T) {
	findings := runUnusedDeclarationAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Private Type ColumnHeader
  name As String
  index As Long
End Type
Private Sub Helper(ByVal columns As Variant)
  Dim column As Variant
  For Each column In columns
    Debug.Print column.index
    Debug.Print column.name
  Next column
End Sub
`})
	if got := findingsByCode(findings, "VBA262"); len(got) != 0 {
		t.Fatalf("VBA262 findings = %+v, want none for late-bound member access", got)
	}
}

func TestUnusedUDTMemberFailsOpenOnVariantEscape(t *testing.T) {
	findings := runUnusedDeclarationAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Private Type TPoint
  X As Long
  Y As Long
End Type
Private Sub Helper()
  Dim p As TPoint
  Dim v As Variant
  v = p
End Sub
`})
	if got := findingsByCode(findings, "VBA262"); len(got) != 0 {
		t.Fatalf("VBA262 findings = %+v, want none after Variant escape", got)
	}
}

// ---------- VBA263/VBA264: definite assignment ----------

func TestNeverAssignedVariableReportsScalarRead(t *testing.T) {
	findings := runUnusedDeclarationAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Public Sub Run()
  Dim total As Long
  Dim other As Long
  other = 2
  Debug.Print total + other
End Sub
`})
	got := findingsByCode(findings, "VBA263")
	if len(got) != 1 || !strings.Contains(got[0].Message, "total") {
		t.Fatalf("VBA263 findings = %+v, want one finding naming total", got)
	}
}

func TestNeverAssignedVariableSkipsByRefMutation(t *testing.T) {
	findings := runUnusedDeclarationAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Private Sub Fill(ByRef target As Long)
  target = 5
End Sub
Public Sub Run()
  Dim total As Long
  Fill total
  Debug.Print total
End Sub
`})
	if got := findingsByCode(findings, "VBA263"); len(got) != 0 {
		t.Fatalf("VBA263 findings = %+v, want none for ByRef-mutated variable", got)
	}
}

func TestUnassignedReadReportsStraightLineRead(t *testing.T) {
	findings := runUnusedDeclarationAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Public Sub Run()
  Dim total As Long
  Dim copy As Long
  copy = total
  total = 1
End Sub
`})
	got := findingsByCode(findings, "VBA264")
	if len(got) != 1 || !strings.Contains(got[0].Message, "total") {
		t.Fatalf("VBA264 findings = %+v, want one finding naming total", got)
	}
}

func TestUnassignedReadReportsPartialBranchAssignment(t *testing.T) {
	findings := runUnusedDeclarationAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Public Sub Run(ByVal flag As Boolean)
  Dim total As Long
  Dim copy As Long
  If flag Then
    total = 1
  End If
  copy = total
End Sub
`})
	got := findingsByCode(findings, "VBA264")
	if len(got) != 1 || !strings.Contains(got[0].Message, "total") {
		t.Fatalf("VBA264 findings = %+v, want one finding naming total", got)
	}
}

func TestUnassignedReadSkipsNextVariableAndErrorSuppression(t *testing.T) {
	findings := runUnusedDeclarationAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Public Sub Run()
On Error Resume Next
  Dim x As Long
  Dim name As String
  name = "a"
  For x = 0 To 3
    If name = "b" Then
    End If
  Next x
End Sub
`})
	if got := findingsByCode(findings, "VBA264"); len(got) != 0 {
		t.Fatalf("VBA264 findings = %+v, want none (Next-variable bookkeeping and On Error flow must not flood)", got)
	}
}

func TestUnassignedReadSkipsVariableAssignedOnAllPaths(t *testing.T) {
	findings := runUnusedDeclarationAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Public Sub Run(ByVal flag As Boolean)
  Dim total As Long
  Dim copy As Long
  If flag Then
    total = 1
  Else
    total = 2
  End If
  copy = total
End Sub
`})
	if got := findingsByCode(findings, "VBA264"); len(got) != 0 {
		t.Fatalf("VBA264 findings = %+v, want none", got)
	}
}
