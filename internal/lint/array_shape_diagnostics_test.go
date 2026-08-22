package lint

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestLintConstantAssignmentsAndInvalidArrayBounds(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Public Const Limit As Long = 2
Public Sub Run()
  Const LocalLimit As Long = Limit + 1
  Dim bad(5 To 2) As Long
  Dim badConst(0 To LocalLimit - 4) As Long
  Dim good(0 To Limit) As Long
  Limit = 3
  LocalLimit = 4
End Sub
`)
	issues, err := (Linter{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB060"); len(got) != 2 {
		t.Fatalf("VB060 = %#v, want 2", got)
	}
	if got := issuesByCode(issues, "VB061"); len(got) != 2 {
		t.Fatalf("VB061 = %#v, want 2", got)
	}
}

func TestLintArrayBoundsRemainConservativeForDynamicExpressions(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal lower As Long, ByVal upper As Long)
  Dim values(lower To upper) As Long
  Dim dynamicValue() As Long
  ReDim dynamicValue(lower To upper)
End Sub
`)
	issues, err := (Linter{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB061"); len(got) != 0 {
		t.Fatalf("dynamic bounds produced VB061: %#v", got)
	}
}

func TestLintArrayShapeResolvesEnumAndShadowingWithOptionBase(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Option Base 1
Public Const Lower As Long = 3
Public Enum Modes
  ModeBad = (4)
End Enum
Public Sub Run()
  Dim Lower As Long
  Dim good(ModeBad To 5) As Long
  Dim bad(0) As Long
  Lower = 2
  ModeBad = 5
End Sub
`)
	issues, err := (Linter{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB060"); len(got) != 1 {
		t.Fatalf("shadowed module Const must remain writable while enum assignment is rejected: %#v", got)
	}
	if got := issuesByCode(issues, "VB061"); len(got) != 1 {
		t.Fatalf("Option Base 1 should make Dim bad(0) impossible: %#v", got)
	}
}

func TestLintArrayBoundsRespectProcedureShadowing(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Public Const Lower As Long = 3
Public Sub Run()
  Dim Lower As Long
  Dim values(Lower To 1) As Long
End Sub
`)
	issues, err := (Linter{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB061"); len(got) != 0 {
		t.Fatalf("local scalar must shadow module Const in array bounds: %#v", got)
	}
}

func TestLintArrayShapeRulesCannotBeInlineSuppressed(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Public Const Limit As Long = 2
Public Sub Run()
  Dim bad(3 To 1) As Long ' xlflow:disable-line VB061
  Limit = 3 ' xlflow:disable-line VB060
End Sub
`)
	result, err := (Linter{RootDir: dir, Config: config.Default()}).RunResult()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(result.Issues, "VB060"); len(got) != 1 {
		t.Fatalf("VB060 was suppressible: %#v", result.Issues)
	}
	if got := issuesByCode(result.Issues, "VB061"); len(got) != 1 {
		t.Fatalf("VB061 was suppressible: %#v", result.Issues)
	}
	if !hasWarning(result.Warnings, "unsupported_inline_suppression_rule", "VB060") || !hasWarning(result.Warnings, "unsupported_inline_suppression_rule", "VB061") {
		t.Fatalf("missing unsupported suppression warnings: %#v", result.Warnings)
	}
}

func TestLintConstantAssignmentUsesProjectVisibleConstants(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeLintModule(t, dir, "Constants.bas", `Option Explicit
Public Const ProjectLimit As Long = 2
`)
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  ProjectLimit = 3
  Constants.ProjectLimit = 4
End Sub
`)
	issues, err := (Linter{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB060"); len(got) != 2 {
		t.Fatalf("project Const assignment was not diagnosed: %#v", got)
	}
}

func TestLintConstantAssignmentDoesNotCrossPrivateModuleBoundary(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeLintModule(t, dir, "PrivateConstants.bas", `Option Explicit
Private Const PrivateLimit As Long = 2
`)
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  PrivateLimit = 3
End Sub
`)
	issues, err := (Linter{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB060"); len(got) != 0 {
		t.Fatalf("private module Const leaked into another module: %#v", got)
	}
}

func TestLintConstantAssignmentRemainsConservativeForAmbiguousProjectNames(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeLintModule(t, dir, "ConstantsA.bas", `Option Explicit
Public Const ProjectLimit As Long = 2
`)
	writeLintModule(t, dir, "ConstantsB.bas", `Option Explicit
Public Const ProjectLimit As Long = 3
`)
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  ProjectLimit = 4
End Sub
`)
	issues, err := (Linter{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB060"); len(got) != 0 {
		t.Fatalf("ambiguous project Const must remain unresolved: %#v", got)
	}
}

func TestLintConstantAssignmentRecognizesBuiltinConstants(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  vbCrLf = "x"
End Sub
`)
	issues, err := (Linter{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB060"); len(got) != 1 {
		t.Fatalf("builtin Const assignment was not diagnosed: %#v", issues)
	}
}

func TestLintConstantAssignmentRecognizesQualifiedConstants(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Public Const Limit As Long = 2
Public Enum Modes
  ModeBad = 1
End Enum
Public Sub Run()
  Main.Limit = 3
  Modes.ModeBad = 2
  VBA.vbCrLf = "x"
End Sub
`)
	issues, err := (Linter{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB060"); len(got) != 3 {
		t.Fatalf("qualified Const assignments = %#v, want 3", got)
	}
}

func TestLintConstantAssignmentRecognizesMeConstants(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	_ = dir
	issues, err := (Linter{RootDir: dir, Config: config.Default(), ModuleKind: "class"}).LintSource("Widget.cls", []byte(`Option Explicit
Private Const Limit As Long = 2
Public Sub Run()
  Me.Limit = 3
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB060"); len(got) != 1 {
		t.Fatalf("Me Const assignment = %#v, want 1", got)
	}
}

func TestLintConstantAssignmentIgnoresWritableExcelPropertyChains(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Option Explicit
Public Const Hidden As Long = 1
Public Enum Modes
  ModeBad = 1
End Enum
Public Sub Test(ByVal ws As Worksheet)
  ws.Rows(1).Hidden = True
  ws.Rows(1).Hidden = False
  ws.Columns(1).Hidden = False
  With ws.Rows(1)
    .Hidden = False
  End With

  Dim receiver As Object
  With receiver
    .Modes.ModeBad = 2
  End With

  Dim rng As Range
  Set rng = ws.Rows(1)
  rng.Hidden = False
End Sub
`
	writeLintModule(t, dir, "Main.bas", source)
	issues, err := (Linter{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB060"); len(got) != 0 {
		t.Fatalf("writable Excel property assignments produced VB060: %#v", got)
	}

	classSource := `Option Explicit
Private Const Limit As Long = 2
Public Sub Run()
  Me.Limit = 3
End Sub
`
	classIssues, err := (Linter{RootDir: dir, Config: config.Default(), ModuleKind: "class"}).LintSource("Widget.cls", []byte(classSource))
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(classIssues, "VB060"); len(got) != 1 || got[0].Symbol != "Me.Limit" {
		t.Fatalf("Me Const assignment = %#v, want one VB060 for Me.Limit", got)
	}
}

func TestLintArrayShapeDoesNotTreatProcedureParametersAsDeclarationBounds(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Private Function Build(ParamArray keyItems() As Variant) As Variant
  Build = Empty
End Function

Private Sub Probe(ByVal lower As Long, ByVal upper As Long)
End Sub
`)
	issues, err := (Linter{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB061"); len(got) != 0 {
		t.Fatalf("procedure parameter lists must not produce VB061: %#v", got)
	}
}
