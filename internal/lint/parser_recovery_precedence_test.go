package lint

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestParserRecoveryIsNotSuppressedByUnrelatedSyntaxIssues(t *testing.T) {
	tests := []struct {
		name   string
		source string
		code   string
	}{
		{
			name: "different line",
			code: "VB032",
			source: `Attribute VB_Name = "Main"
Option Explicit
Sub Main()
    ?? "bad"
    Debug.Print (1
End Sub
`,
		},
		{
			name: "colon separated statement",
			code: "VB032",
			source: `Attribute VB_Name = "Main"
Option Explicit
Sub Main()
    ?? "bad" : Debug.Print (1
End Sub
`,
		},
		{
			name: "colon separated lexical and parser defects",
			code: "VB008",
			source: `Attribute VB_Name = "Main"
Option Explicit
Sub Main()
    Debug.Print “bad” : Debug.Print (1
End Sub
`,
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			issues, err := (Linter{Config: config.Default()}).LintSource("Main.bas", []byte(tc.source))
			if err != nil {
				t.Fatal(err)
			}
			if got := issuesByCode(issues, tc.code); len(got) != 1 {
				t.Fatalf("%s = %+v, want one finding; issues = %+v", tc.code, got, issues)
			}
			if got := issuesByCode(issues, "VB014"); len(got) != 1 {
				t.Fatalf("unrelated parser recovery was suppressed: issues = %+v", issues)
			}
		})
	}
}

func TestParserRecoveryIsNotSuppressedByUnrelatedVB059OnSameLine(t *testing.T) {
	source := `Attribute VB_Name = "Main"
Option Explicit
Sub Probe()
    Call Foo + Bar : Debug.Print (1
End Sub
`
	issues, err := (Linter{Config: config.Default()}).LintSource("Main.bas", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, callSyntaxDiagnosticCode); len(got) != 1 {
		t.Fatalf("VB059 = %+v, want one finding; issues = %+v", got, issues)
	}
	if got := issuesByCode(issues, "VB014"); len(got) != 1 {
		t.Fatalf("unrelated parser recovery was suppressed by same-line VB059: issues = %+v", issues)
	}
}

func TestParserRecoveryIsSuppressedByOverlappingVB059(t *testing.T) {
	source := `Attribute VB_Name = "Main"
Option Explicit
Sub Probe()
    Call Foo + Bar
End Sub
`
	issues, err := (Linter{Config: config.Default()}).LintSource("Main.bas", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, callSyntaxDiagnosticCode); len(got) != 1 {
		t.Fatalf("VB059 = %+v, want one finding; issues = %+v", got, issues)
	}
	if got := issuesByCode(issues, "VB014"); len(got) != 0 {
		t.Fatalf("overlapping VB059 should suppress duplicate VB014: %+v", got)
	}
}

func TestParserRecoveryRetainsMultipleUnclosedBlockDiagnostics(t *testing.T) {
	source := `Attribute VB_Name = "Main"
Option Explicit
Sub Main()
    If enabled Then
        For Each item In items
            Debug.Print item
End Sub
`
	issues, err := (Linter{Config: config.Default()}).LintSource("Main.bas", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	got := issuesByCode(issues, "VB014")
	if len(got) != 2 {
		t.Fatalf("unclosed blocks = %+v, want one VB014 per missing closer", got)
	}
	kinds := map[string]bool{}
	for _, issue := range got {
		kinds[issue.BlockKind] = true
	}
	if !kinds["if"] || !kinds["for"] {
		t.Fatalf("unclosed block kinds = %#v, want if and for", kinds)
	}
}
