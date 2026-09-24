package analyze

import (
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/typedb"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
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

func TestUnusedParameterSkipsTestPrefixedWithEventsCallback(t *testing.T) {
	findings := runUnusedDeclarationAnalysis(t, map[string]string{"Handler.cls": `Option Explicit
Private WithEvents TestApp As Application
Private Sub TestApp_WindowBeforeDoubleClick(ByVal Sel As Selection, ByRef Cancel As Boolean)
End Sub
`})
	if got := findingsByCode(findings, "VBA260"); len(got) != 0 {
		t.Fatalf("VBA260 findings = %+v, want none for a Test-prefixed WithEvents callback", got)
	}
}

func TestUnusedParameterReportsOneBasedParameterRange(t *testing.T) {
	source := `Option Explicit
Private Sub Helper(ByVal extra As Long)
End Sub
`
	findings := runUnusedDeclarationAnalysis(t, map[string]string{"Main.bas": source})
	got := findingsByCode(findings, "VBA260")
	if len(got) != 1 {
		t.Fatalf("VBA260 findings = %+v, want one", got)
	}
	line := strings.Split(source, "\n")[got[0].Line-1]
	if got[0].Column < 1 || got[0].EndColumn > len(line)+1 {
		t.Fatalf("VBA260 range %d-%d outside line %q", got[0].Column, got[0].EndColumn, line)
	}
	if fragment := line[got[0].Column-1 : got[0].EndColumn-1]; fragment != "ByVal extra As Long" {
		t.Fatalf("VBA260 range %d-%d covers %q, want the parameter declarator", got[0].Column, got[0].EndColumn, fragment)
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

func TestUnusedParameterSkipsQualifiedImplementsMembers(t *testing.T) {
	findings := runUnusedDeclarationAnalysis(t, map[string]string{
		"IFace.cls": `Option Explicit
Public Sub DoWork(ByVal arg As Long)
End Sub
`,
		"Impl.cls": `Option Explicit
Implements Lib.IFace
Private Sub IFace_DoWork(ByVal arg As Long)
End Sub
`,
	})
	if got := findingsByCode(findings, "VBA260"); len(got) != 0 {
		t.Fatalf("VBA260 findings = %+v, want none for a qualified Implements member", got)
	}
}

func TestUnusedParameterSkipsCrossModuleDynamicEntry(t *testing.T) {
	findings := runUnusedDeclarationAnalysis(t, map[string]string{
		"Main.bas": `Option Explicit
Private Sub Helper(ByVal arg As Long)
End Sub
`,
		"Driver.bas": `Option Explicit
Public Sub Run()
  Application.Run "Main.Helper", 1
End Sub
`,
	})
	if got := findingsByCode(findings, "VBA260"); len(got) != 0 {
		t.Fatalf("VBA260 findings = %+v, want none for a cross-module dynamic entry point", got)
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

func TestUnusedPrivateConstCountsQualifiedReferenceUnderShadow(t *testing.T) {
	findings := runUnusedDeclarationAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Private Const Limit As Long = 10
Public Sub Run()
  Dim Limit As Long
  Limit = 3
  Debug.Print Main.Limit
End Sub
`})
	if got := findingsByCode(findings, "VBA261"); len(got) != 0 {
		t.Fatalf("VBA261 findings = %+v, want none for a qualified module reference", got)
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

func TestUnusedUDTMemberSkipsShadowedModuleVariable(t *testing.T) {
	// A local of another type shadows the module-level UDT variable; member
	// expressions on it are late-bound and must not resolve against the hidden
	// module variable's private type.
	findings := runUnusedDeclarationAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Private Type TPoint
  X As Long
End Type
Private Type TMeta
  Tag As String
End Type
Private p As TPoint
Private Sub Helper()
  Dim p As Object
  Debug.Print p.Tag
End Sub
`})
	for _, finding := range findingsByCode(findings, "VBA262") {
		if strings.Contains(finding.Message, "Tag") {
			t.Fatalf("VBA262 findings = %+v, want no finding for Tag through the shadowed receiver", findingsByCode(findings, "VBA262"))
		}
	}
}

func TestUnusedUDTMemberEscapesThroughNestedArgument(t *testing.T) {
	// The writable argument root is the parenthesized expression; the UDT
	// variable access sits one level below and must still escape the type.
	findings := runUnusedDeclarationAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Private Type TPoint
  X As Long
End Type
Private Sub Helper()
  Dim p As TPoint
  UnknownMutate (p)
End Sub
`})
	if got := findingsByCode(findings, "VBA262"); len(got) != 0 {
		t.Fatalf("VBA262 findings = %+v, want none after a nested writable argument escapes the type", got)
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

func TestAssignmentAmbiguousTargetIsolation(t *testing.T) {
	cases := []struct {
		text   string
		target string
		shaped bool
	}{
		{`Mid$(s, i, 1) = "x"`, "s", true},
		{`Mid(s, 1) = "x"`, "s", true},
		{`LSet lval = "pad"`, "lval", true},
		{`RSet rval = "pad"`, "rval", true},
		// Mid$ as a function call on the right-hand side is not ambiguous.
		{`x = Mid$(s, 1, 1)`, "", false},
		{`Debug.Print Mid$(s, 1, 1)`, "", false},
		// Ordinary assignments keep normal read/write modeling.
		{`total = 1`, "", false},
		{`Midpoint = 1`, "", false},
	}
	for _, tc := range cases {
		statement := procedureir.Statement{Kind: procedureir.StatementAssignment, Text: tc.text}
		target, shaped := assignmentAmbiguousTarget(statement)
		if shaped != tc.shaped || target != tc.target {
			t.Fatalf("assignmentAmbiguousTarget(%q) = (%q, %v), want (%q, %v)", tc.text, target, shaped, tc.target, tc.shaped)
		}
	}
}

func TestAssignmentAnalysisSuppressesMidStatementTarget(t *testing.T) {
	findings := runUnusedDeclarationAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Public Sub Run()
  Dim s As String
  Dim i As Long
  Dim t As String
  Mid$(s, i, 1) = "x"
  t = s
End Sub
`})
	for _, f := range findingsByCode(findings, "VBA263") {
		if strings.Contains(f.Message, " s ") {
			t.Fatalf("VBA263 reported Mid$ write target: %+v", f)
		}
	}
	for _, f := range findingsByCode(findings, "VBA264") {
		if strings.Contains(f.Message, " s ") {
			t.Fatalf("VBA264 reported Mid$ write target: %+v", f)
		}
	}
}

func TestAssignmentAnalysisSuppressesLSetAndRSetTargets(t *testing.T) {
	findings := runUnusedDeclarationAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Public Sub Run()
  Dim lval As String
  Dim rval As String
  Dim t As String
  LSet lval = "pad"
  RSet rval = "pad"
  t = lval
  t = rval
End Sub
`})
	if got := findingsByCode(findings, "VBA263"); len(got) != 0 {
		t.Fatalf("VBA263 findings = %+v, want none for LSet/RSet write targets", got)
	}
	if got := findingsByCode(findings, "VBA264"); len(got) != 0 {
		t.Fatalf("VBA264 findings = %+v, want none for LSet/RSet write targets", got)
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
