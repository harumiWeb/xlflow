package lint

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestConditionalBranchSyntaxIssuesReportsMissingThen(t *testing.T) {
	source := `Option Explicit
Sub Main()
  If ready
    Debug.Print "x"
  End If
End Sub
`
	issues := (Linter{RootDir: "."}).conditionalBranchSyntaxIssues("Main.bas", source)
	if len(issues) != 1 {
		t.Fatalf("issues = %#v, want one missing-Then finding", issues)
	}
	issue := issues[0]
	if issue.Code != conditionalBranchDiagnosticCode || issue.Kind != "missing_then" || issue.Symbol != "If" {
		t.Fatalf("unexpected missing-Then issue: %+v", issue)
	}
	if issue.Line != 3 || issue.Column != 3 || issue.EndLine != 3 || issue.EndColumn <= issue.Column {
		t.Fatalf("missing-Then range = %+v, want statement range on line 3", issue)
	}
}

func TestConditionalBranchSyntaxIssuesReportsElsePlacement(t *testing.T) {
	tests := []struct {
		name string
		src  string
		line int
		kind string
	}{
		{
			name: "orphan Else",
			src:  "Sub Main()\n  Else\n    Debug.Print \"x\"\nEnd Sub\n",
			line: 2,
			kind: "else_without_if",
		},
		{
			name: "orphan ElseIf",
			src:  "Sub Main()\n  ElseIf fallback Then\n    Debug.Print \"x\"\nEnd Sub\n",
			line: 2,
			kind: "elseif_without_if",
		},
		{
			name: "ElseIf after Else",
			src:  "Sub Main()\n  If ready Then\n    Debug.Print \"x\"\n  Else\n    Debug.Print \"y\"\n  ElseIf fallback Then\n    Debug.Print \"z\"\n  End If\nEnd Sub\n",
			line: 6,
			kind: "elseif_after_else",
		},
		{
			name: "duplicate Else",
			src:  "Sub Main()\n  If ready Then\n    Debug.Print \"x\"\n  Else\n    Debug.Print \"y\"\n  Else\n    Debug.Print \"z\"\n  End If\nEnd Sub\n",
			line: 6,
			kind: "duplicate_else",
		},
	}

	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			issues := (Linter{RootDir: "."}).conditionalBranchSyntaxIssues("Main.bas", test.src)
			if len(issues) != 1 {
				t.Fatalf("issues = %#v, want one VB062 finding", issues)
			}
			if issues[0].Code != conditionalBranchDiagnosticCode || issues[0].Kind != test.kind || issues[0].Line != test.line {
				t.Fatalf("issue = %+v, want kind %q at line %d", issues[0], test.kind, test.line)
			}
		})
	}
}

func TestConditionalBranchSyntaxIssuesDoesNotCascadeMissingThenIntoElse(t *testing.T) {
	source := "Sub Main()\n  If ready\n  Else\n  End If\nEnd Sub\n"
	issues := (Linter{RootDir: "."}).conditionalBranchSyntaxIssues("Main.bas", source)
	if len(issues) != 1 || issues[0].Kind != "missing_then" {
		t.Fatalf("issues = %#v, want only missing_then for one malformed If block", issues)
	}
}

func TestConditionalBranchSyntaxIssuesRemainsQuietForValidAndAmbiguousForms(t *testing.T) {
	valid := `Sub Main()
  If value = "Then" Then
    Debug.Print "x"
  ElseIf fallback Then
    Debug.Print "y"
  Else
    Debug.Print "z"
  End If
  If amount# Then
    Debug.Print amount
  End If
  If stamp = #1/1/2020# Then
    Debug.Print stamp
  End If
End Sub
`
	if issues := (Linter{RootDir: "."}).conditionalBranchSyntaxIssues("Main.bas", valid); len(issues) != 0 {
		t.Fatalf("valid branch chain produced findings: %+v", issues)
	}

	// Conditional-compilation directives make ownership unknowable from an
	// exported source buffer; generic VB014 remains the caller's fallback.
	ambiguous := `Sub Main()
#If Win64 Then
  If ready Then
#End If
  Else
    Debug.Print "x"
  End If
End Sub
`
	if issues := (Linter{RootDir: "."}).conditionalBranchSyntaxIssues("Main.bas", ambiguous); len(issues) != 0 {
		t.Fatalf("conditional-compilation source should fail open: %+v", issues)
	}
}

func TestConditionalBranchSyntaxIssuesHandlesColonAndLineNumbers(t *testing.T) {
	remComment := "Sub Main()\n  Rem: Else\nEnd Sub\n"
	if issues := (Linter{RootDir: "."}).conditionalBranchSyntaxIssues("Main.bas", remComment); len(issues) != 0 {
		t.Fatalf("Rem comment colon should consume branch-looking text: %+v", issues)
	}

	source := `Sub Main()
10 If ready Then: Debug.Print "x": Else
20 Debug.Print "y"
30 End If
End Sub
`
	if issues := (Linter{RootDir: "."}).conditionalBranchSyntaxIssues("Main.bas", source); len(issues) != 0 {
		t.Fatalf("one-line conditional with colon branches should remain quiet: %+v", issues)
	}

	missing := "Sub Main()\n10 If ready\nEnd Sub\n"
	issues := (Linter{RootDir: "."}).conditionalBranchSyntaxIssues("Main.bas", missing)
	if len(issues) != 1 || issues[0].Kind != "missing_then" || issues[0].Line != 2 || issues[0].Column != 4 {
		t.Fatalf("line-numbered missing Then = %+v", issues)
	}
}

func TestLinterPublishesConditionalBranchSyntaxDiagnostic(t *testing.T) {
	source := []byte("Sub Main()\n  If ready\n    Debug.Print \"x\"\n  End If\nEnd Sub\n")
	issues, err := (Linter{Config: config.Default()}).LintSource("Main.bas", source)
	if err != nil {
		t.Fatal(err)
	}
	got := issuesByCode(issues, conditionalBranchDiagnosticCode)
	if len(got) != 1 {
		t.Fatalf("VB062 = %+v, want one diagnostic; issues = %+v", got, issues)
	}
	if got[0].Kind != "missing_then" || got[0].Line != 2 {
		t.Fatalf("VB062 = %+v, want missing Then on line 2", got[0])
	}
	if recovery := issuesByCode(issues, "VB014"); len(recovery) != 0 {
		t.Fatalf("specific conditional syntax should not add VB014: %+v", recovery)
	}
}

func TestLinterDoesNotClassifyModuleScopeConditionalSyntaxAsVB062(t *testing.T) {
	source := []byte(`Attribute VB_Name = "Main"
Option Explicit
If ready Then
  Debug.Print "x"
Else
  Debug.Print "y"
End If
`)
	issues, err := (Linter{Config: config.Default()}).LintSource("Main.bas", source)
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, conditionalBranchDiagnosticCode); len(got) != 0 {
		t.Fatalf("module-scope conditional syntax should not produce VB062: %+v", got)
	}
	if got := issuesByCode(issues, "VB014"); len(got) == 0 {
		t.Fatalf("module-scope conditional syntax should retain existing recovery/placement diagnostics: %+v", issues)
	}
}

func TestConditionalBranchDoesNotHideMalformedNeighboringRecovery(t *testing.T) {
	source := []byte("Sub Main()\n  If ready\n    Debug.Print (1\nEnd Sub\n")
	issues, err := (Linter{Config: config.Default()}).LintSource("Main.bas", source)
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, conditionalBranchDiagnosticCode); len(got) != 1 {
		t.Fatalf("VB062 = %+v, want one missing-Then finding; all issues = %+v", got, issues)
	}
	if got := issuesByCode(issues, "VB014"); len(got) == 0 {
		t.Fatalf("malformed neighboring expression should retain VB014: %+v", issues)
	}
}

func TestSpecificSyntaxDoesNotHideUnclosedBlockRecovery(t *testing.T) {
	source := []byte("Attribute VB_Name = \"Main\"\nOption Explicit\nSub Main()\n  TypeOf value Is Collection Extra\n  If ready Then\n    Debug.Print \"x\"\nEnd Sub\n")
	issues, err := (Linter{Config: config.Default()}).LintSource("Main.bas", source)
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, typeOfSyntaxDiagnosticCode); len(got) != 1 {
		t.Fatalf("VB065 = %+v, want one TypeOf finding; all issues = %+v", got, issues)
	}
	if got := issuesByCode(issues, "VB014"); len(got) == 0 {
		t.Fatalf("unclosed neighboring block should retain VB014: %+v", issues)
	}
}
