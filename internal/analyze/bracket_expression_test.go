package analyze

import (
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vba/sourceproject"
)

func TestBracketExpressionFindingsReportsBareHostExpressions(t *testing.T) {
	const source = `Option Explicit
Public Enum Options
  [_enumLast] = 1
End Enum
Public Sub Run()
  result = [A1]
  result = Foo([A2])
  result = Me.[Member]
  result = values(index)
  Dim [Declared] As String
  result = [Declared]
  result = [_enumLast]
  result = "literal [A4]"
  ' [A5]
End Sub
`
	root := t.TempDir()
	ir, lines := buildBracketExpressionTestIR(t, source)
	file := parsedFile{Path: filepath.Join(root, "Main.bas"), Lines: lines, Module: "Main", IR: ir}
	proc := sourceProceduresFromIRRef(&file.IR)[0]

	findings := bracketExpressionFindings(root, file, proc, nil)
	if len(findings) != 2 {
		t.Fatalf("VBA262 findings = %+v, want two bare bracket expressions", findings)
	}
	if findings[0].Line != 6 || findings[0].Column < 1 || findings[0].File != "Main.bas" {
		t.Fatalf("first VBA262 finding = %+v, want [A1] location in Main.bas", findings[0])
	}
	if findings[1].Line != 7 || findings[1].ScopeEndLine != proc.EndLine {
		t.Fatalf("second VBA262 finding = %+v, want Foo([A2]) location scoped to procedure", findings[1])
	}
}

func TestBracketExpressionCandidatesExcludeMembersAndDeclarations(t *testing.T) {
	const source = `Public Sub Run()
  result = Me.[Member]
  Dim [Declared] As String
  result = values(index)
  result = "[Literal]"
  ' [Comment]
End Sub
`
	ir, _ := buildBracketExpressionTestIR(t, source)
	proc := sourceProceduresFromIRRef(&ir)[0]
	if got := bracketExpressionCandidates(parsedFile{}, proc, nil); len(got) != 0 {
		t.Fatalf("bracket expression candidates = %+v, want none", got)
	}
}

func TestBracketExpressionFindingsExcludeProjectDeclaration(t *testing.T) {
	cfg := config.Default()
	cfg.Analyze.DetectHostBracketExpressions = true
	project := sourceproject.SourceProject{Files: []sourceproject.SourceFile{
		{
			Path:       "virtual/Main.bas",
			ModuleKind: sourceproject.ModuleKindStandard,
			Source:     []byte("Public Sub Run()\n  Debug.Print [SharedValue]\n  Debug.Print [SharedFunction]\n  Debug.Print [Hidden]\n  Debug.Print [A1]\nEnd Sub\n"),
		},
		{
			Path:       "virtual/Shared.bas",
			ModuleKind: sourceproject.ModuleKindStandard,
			Source:     []byte("Public Const [SharedValue] As Long = 1\nPublic Function [SharedFunction]() As Long\nEnd Function\nPrivate Function [Hidden]() As Long\nEnd Function\n"),
		},
	}}
	result, err := (Analyzer{Config: cfg}).AnalyzeProject(t.Context(), project)
	if err != nil {
		t.Fatal(err)
	}
	findings := findingsByCode(result.Findings, "VBA262")
	if len(findings) != 2 || findings[0].Line != 4 || findings[1].Line != 5 {
		t.Fatalf("VBA262 findings = %+v, want inaccessible [Hidden] and unresolved [A1]", findings)
	}
}

func buildBracketExpressionTestIR(t *testing.T, source string) (procedureir.DocumentIR, []string) {
	t.Helper()
	document, err := vbaast.ParseDocument("Main.bas", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	defer document.Close()
	ir, err := procedureir.BuildParsedContext(t.Context(), procedureir.BuildOptions{
		Path:       "Main.bas",
		ModuleName: "Main",
		ModuleKind: "standard",
	}, document)
	if err != nil {
		t.Fatal(err)
	}
	if len(ir.Procedures) != 1 {
		t.Fatalf("procedure count = %d, want one", len(ir.Procedures))
	}
	return ir, normalizedSourceLines(source)
}
