package lint

import (
	"context"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestDeclarationDiagnosticsDuplicateAndPlacement(t *testing.T) {
	tests := []struct {
		name   string
		source string
		want   map[string]int
	}{
		{
			name:   "duplicate locals and options",
			source: "Option Explicit\nOption Explicit\nDim Value As Long\nDim value As String\n",
			want:   map[string]int{"VB046": 2},
		},
		{
			name:   "module declarations share scope",
			source: "Dim PublicName As Long\nSub publicname()\nEnd Sub\n",
			want:   map[string]int{"VB046": 1},
		},
		{
			name:   "module declaration after procedure",
			source: "Option Explicit\nSub First()\nEnd Sub\nPrivate LateValue As Long\n",
			want:   map[string]int{"VB047": 1},
		},
		{
			name:   "option inside procedure",
			source: "Attribute VB_Name = \"Main\"\nSub First()\nOption Explicit\nEnd Sub\n",
			want:   map[string]int{"VB047": 1},
		},
		{
			name:   "legal separate procedures and property accessors",
			source: "Option Explicit\nSub First()\n    Dim value As Long\nEnd Sub\nSub Second()\n    Dim value As String\nEnd Sub\nProperty Get Caption() As String\nEnd Property\nProperty Let Caption(ByVal value As String)\nEnd Property\n",
			want:   map[string]int{},
		},
		{
			name:   "implements may precede options in a class module",
			source: "Implements ITest\nOption Explicit\nSub Run()\nEnd Sub\n",
			want:   map[string]int{},
		},
		{
			name:   "enum and type members",
			source: "Option Explicit\nEnum Status\n    Ready\n    READY\nEnd Enum\nType Payload\n    Value As Long\n    value As Long\nEnd Type\n",
			want:   map[string]int{"VB046": 2},
		},
		{
			name:   "repeated property accessor",
			source: "Option Explicit\nProperty Get Caption() As String\nEnd Property\nProperty Get Caption() As String\nEnd Property\n",
			want:   map[string]int{"VB046": 1},
		},
		{
			name:   "conditional property accessor branches",
			source: "Option Explicit\n#If VBA7 Then\nProperty Let UserData(ByVal value As LongPtr)\n#Else\nProperty Let UserData(ByVal value As Long)\n#End If\nEnd Property\n",
			want:   map[string]int{},
		},
		{
			name:   "parameters locals static and const share procedure scope",
			source: "Option Explicit\nSub Run(ByVal value As Long)\n    Dim VALUE As Long\n    Static other As Long\n    Const OTHER As Long = 1\n    Dim other As Long\nEnd Sub\n",
			want:   map[string]int{"VB046": 3},
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			issues, err := (Linter{}).LintSourceContext(context.Background(), "Main.bas", []byte(tc.source))
			if err != nil {
				t.Fatal(err)
			}
			got := map[string]int{}
			for _, issue := range issues {
				if issue.Code == "VB046" || issue.Code == "VB047" {
					got[issue.Code]++
				}
			}
			for code, want := range tc.want {
				if got[code] != want {
					t.Fatalf("%s count = %d, want %d; issues=%+v", code, got[code], want, issues)
				}
			}
			for code, count := range got {
				if count != tc.want[code] {
					t.Fatalf("unexpected %s count = %d, want %d; issues=%+v", code, count, tc.want[code], issues)
				}
			}
		})
	}
}

func TestDeclarationDiagnosticsFailOpenOnParserRecovery(t *testing.T) {
	source := "Option Explicit\nSub Run()\n    Dim value As Long\n    Dim VALUE As String\n"
	issues, err := (Linter{}).LintSource("Main.bas", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB046"); len(got) != 0 {
		t.Fatalf("parser recovery must not produce VB046: %+v", got)
	}
	if got := issuesByCode(issues, "VB047"); len(got) != 0 {
		t.Fatalf("parser recovery must not produce VB047: %+v", got)
	}
}

func TestDeclarationDiagnosticsPreflightFindingsCannotBeSuppressed(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Sub Run()
    Dim value As Long
    ' xlflow:disable-next-line VB046
    Dim VALUE As String
End Sub
`)
	result, err := (Linter{RootDir: dir, Config: config.Default()}).RunResult()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(result.Issues, "VB046"); len(got) != 1 {
		t.Fatalf("VB046 should remain unsuppressed: issues=%+v warnings=%+v", result.Issues, result.Warnings)
	}
	if !hasWarning(result.Warnings, "unsupported_inline_suppression_rule", "VB046") {
		t.Fatalf("expected unsupported suppression warning, got %+v", result.Warnings)
	}
}

func TestDeclarationDiagnosticsIgnoreMutuallyExclusiveBranches(t *testing.T) {
	source := "Option Explicit\n#If DEBUG Then\nDim value As Long\n#Else\nDim VALUE As String\n#End If\n"
	issues, err := (Linter{}).LintSource("Main.bas", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	for _, issue := range issues {
		if issue.Code == "VB046" || issue.Code == "VB047" {
			t.Fatalf("conditional branch produced declaration diagnostic: %+v", issue)
		}
	}
}

func TestDeclarationDiagnosticsCompareRepeatedConditionalBranch(t *testing.T) {
	source := "Option Explicit\n#If DEBUG Then\nDim value As Long\n#End If\n#If DEBUG Then\nDim VALUE As String\n#End If\n"
	issues, err := (Linter{}).LintSource("Main.bas", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB046"); len(got) != 1 {
		t.Fatalf("repeated active conditional branch should report one duplicate: %+v", got)
	}
}

func TestDeclarationDiagnosticsIgnoreStaticallyFalseBranch(t *testing.T) {
	source := "Option Explicit\nEnum API\n    S_OK = 0\n#If False Then\n    Dim S_OK\n#End If\nEnd Enum\n"
	issues, err := (Linter{}).LintSource("Main.bas", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	for _, issue := range issues {
		if issue.Code == "VB046" || issue.Code == "VB047" {
			t.Fatalf("statically false branch produced declaration diagnostic: %+v", issue)
		}
	}
}
