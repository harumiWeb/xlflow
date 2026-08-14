package lint

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestLinterReportsOpenMissingMode(t *testing.T) {
	source := []byte("Sub Main()\n    Open path For As #1\nEnd Sub\n")
	issues, err := (Linter{Config: config.Default()}).LintSource("Main.bas", source)
	if err != nil {
		t.Fatal(err)
	}
	got := issuesByCode(issues, openSyntaxDiagnosticCode)
	if len(got) != 1 || got[0].Kind != "missing_mode" {
		t.Fatalf("VB064 = %+v, want one missing-mode diagnostic; all issues = %+v", got, issues)
	}
	if recovery := issuesByCode(issues, "VB014"); len(recovery) != 0 {
		t.Fatalf("specific Open recovery should suppress overlapping VB014: %+v", recovery)
	}
}

func TestLinterReportsOpenMissingModeInColonSeparatedStatement(t *testing.T) {
	source := []byte("Sub Main()\n    Open path For As #1: Debug.Print \"after\"\nEnd Sub\n")
	issues, err := (Linter{Config: config.Default()}).LintSource("Main.bas", source)
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, openSyntaxDiagnosticCode); len(got) != 1 || got[0].Kind != "missing_mode" {
		t.Fatalf("VB064 = %+v, want one colon-separated missing-mode diagnostic", got)
	}
	if got := issuesByCode(issues, "VB014"); len(got) != 0 {
		t.Fatalf("colon-separated Open recovery should suppress overlapping VB014: %+v", got)
	}
}

func TestLinterReportsTypeOfTrailingToken(t *testing.T) {
	source := []byte("Sub Main()\n    Dim value As Object\n    TypeOf value Is Collection Extra\nEnd Sub\n")
	issues, err := (Linter{Config: config.Default()}).LintSource("Main.bas", source)
	if err != nil {
		t.Fatal(err)
	}
	got := issuesByCode(issues, typeOfSyntaxDiagnosticCode)
	if len(got) != 1 || got[0].Kind != "trailing_token" {
		t.Fatalf("VB065 = %+v, want one trailing-token diagnostic; all issues = %+v", got, issues)
	}
	if recovery := issuesByCode(issues, "VB014"); len(recovery) != 0 {
		t.Fatalf("specific TypeOf recovery should not add VB014: %+v", recovery)
	}
}

func TestLinterKeepsAmbiguousOpenAndTypeOfRecoveryAsVB014(t *testing.T) {
	tests := []struct {
		name   string
		source string
	}{
		{
			name:   "missing file number",
			source: "Sub Main()\n    Open path For Input As\nEnd Sub\n",
		},
		{
			name:   "missing TypeOf Is",
			source: "Sub Main()\n    TypeOf value Collection\nEnd Sub\n",
		},
		{
			name:   "TypeOf with malformed neighboring recovery",
			source: "Sub Main()\n    TypeOf value Is Collection Extra: Debug.Print (1\nEnd Sub\n",
		},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			issues, err := (Linter{Config: config.Default()}).LintSource("Main.bas", []byte(test.source))
			if err != nil {
				t.Fatal(err)
			}
			if got := issuesByCode(issues, openSyntaxDiagnosticCode); len(got) != 0 {
				t.Fatalf("VB064 = %+v, want generic fallback", got)
			}
			if got := issuesByCode(issues, typeOfSyntaxDiagnosticCode); len(got) != 0 {
				t.Fatalf("VB065 = %+v, want generic fallback", got)
			}
			if got := issuesByCode(issues, "VB014"); len(got) == 0 {
				t.Fatalf("issues = %+v, want VB014 fallback", issues)
			}
		})
	}
}

func TestLinterAcceptsValidOpenAndTypeOfForms(t *testing.T) {
	source := []byte("Sub Main()\n    Dim value As Object\n    Open path For Input As #1\n    If TypeOf value Is Collection Then: Debug.Print \"ok\"\nEnd Sub\n")
	issues, err := (Linter{Config: config.Default()}).LintSource("Main.bas", source)
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, openSyntaxDiagnosticCode); len(got) != 0 {
		t.Fatalf("valid Open produced VB064: %+v", got)
	}
	if got := issuesByCode(issues, typeOfSyntaxDiagnosticCode); len(got) != 0 {
		t.Fatalf("valid TypeOf produced VB065: %+v", got)
	}
}

func TestLinterAcceptsNestedTypeOfExpression(t *testing.T) {
	source := []byte("Sub Main()\n    Dim value As Object\n    If (TypeOf value Is Collection) Then\n        Debug.Print \"ok\"\n    End If\nEnd Sub\n")
	issues, err := (Linter{Config: config.Default()}).LintSource("Main.bas", source)
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, typeOfSyntaxDiagnosticCode); len(got) != 0 {
		t.Fatalf("nested valid TypeOf produced VB065: %+v", got)
	}
}
