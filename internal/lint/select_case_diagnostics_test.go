package lint

import (
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
)

func TestSelectCaseSyntaxIssues(t *testing.T) {
	t.Parallel()
	tests := []struct {
		name   string
		source string
		want   []selectCaseExpectation
	}{
		{
			name: "case outside select",
			source: `Sub Probe()
    Case 1
End Sub
`,
			want: []selectCaseExpectation{{kind: "case_outside_select", line: 2}},
		},
		{
			name: "case else outside select",
			source: `Sub Probe()
    Case Else
End Sub
`,
			want: []selectCaseExpectation{{kind: "case_outside_select", line: 2}},
		},
		{
			name: "duplicate case else",
			source: `Sub Probe()
    Select Case value
    Case 1
        Debug.Print "one"
    Case Else
        Debug.Print "default"
    Case Else
        Debug.Print "duplicate"
    End Select
End Sub
`,
			want: []selectCaseExpectation{{kind: "duplicate_case_else", line: 7}},
		},
		{
			name: "case after case else",
			source: `Sub Probe()
    Select Case value
    Case Else
        Debug.Print "default"
    Case 1
        Debug.Print "late"
    End Select
End Sub
`,
			want: []selectCaseExpectation{{kind: "case_after_else", line: 5}},
		},
		{
			name: "nested select blocks",
			source: `Sub Probe()
    Select Case outerValue
    Case 1
        Select Case innerValue
        Case 2
            Debug.Print "two"
        Case Else
            Debug.Print "inner default"
        End Select
    Case Else
        Debug.Print "outer default"
    End Select
End Sub
`,
		},
		{
			name: "colon separated case bodies",
			source: `Sub Probe()
    Select Case value
    Case 1: Debug.Print "one": Debug.Print "again"
    Case Else: Debug.Print "default"
    End Select
End Sub
`,
		},
		{
			name: "colon inside date literal",
			source: `Sub Probe()
    Select Case #12:00:00 PM#
    Case 1: Debug.Print "one"
    Case Else: Debug.Print "default"
    End Select
End Sub
`,
		},
		{
			name: "colon separated nested select",
			source: `Sub Probe()
    Select Case outerValue
    Case 1: Select Case innerValue
        Case 2: Debug.Print "two"
        Case Else: Debug.Print "inner default"
    End Select
    Case Else: Debug.Print "outer default"
    End Select
End Sub
`,
		},
		{
			name: "words in strings and comments",
			source: `Sub Probe()
    Debug.Print "Case Else: Case 1"
    ' Case 2
    Rem: Case 3
End Sub
`,
		},
		{
			name: "malformed case falls back to parser recovery",
			source: `Sub Probe()
    Select Case value
    Case
        Debug.Print "ambiguous"
    End Select
End Sub
`,
		},
		{
			name: "malformed neighboring statement remains ambiguous",
			source: `Sub Probe()
    Select Case value
    Case 1
        If value Then
    Case Else
        Debug.Print "ambiguous"
    End Select
End Sub
`,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			doc, err := vbaast.ParseDocument("Main.bas", []byte(tt.source))
			if err != nil {
				t.Fatal(err)
			}
			defer doc.Close()
			var issues []selectCaseSyntaxIssue
			err = doc.Read(func(view vbaast.ParsedView) error {
				issues = selectCaseSyntaxIssues(string(view.Source), view.Root)
				return nil
			})
			if err != nil {
				t.Fatal(err)
			}
			if len(issues) != len(tt.want) {
				t.Fatalf("VB063 issues = %#v, want %#v", issues, tt.want)
			}
			for i, want := range tt.want {
				got := issues[i]
				if got.Code != selectCaseDiagnosticCode {
					t.Errorf("issue[%d] code = %q, want %s", i, got.Code, selectCaseDiagnosticCode)
				}
				if got.Kind != want.kind || got.Range.StartLine != want.line {
					t.Errorf("issue[%d] = %#v, want kind=%q line=%d", i, got, want.kind, want.line)
				}
				if got.Message == "" || got.Suggestion == "" {
					t.Errorf("issue[%d] missing actionable text: %#v", i, got)
				}
			}
		})
	}
}

func TestSelectCaseSyntaxIssuesRequireUnambiguousRoot(t *testing.T) {
	t.Parallel()
	if got := selectCaseSyntaxIssues("Case 1\n", nil); got != nil {
		t.Fatalf("nil root produced issues: %#v", got)
	}

	doc, err := vbaast.ParseDocument("Main.bas", []byte("Sub Probe()\n    Select Case value\n    Case\nEnd Sub\n"))
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	err = doc.Read(func(view vbaast.ParsedView) error {
		if got := selectCaseSyntaxIssues(string(view.Source), view.Root); len(got) != 0 {
			t.Fatalf("parser recovery produced specific issues: %#v", got)
		}
		return nil
	})
	if err != nil {
		t.Fatal(err)
	}
}

func TestSelectCaseSyntaxIssueRangesUseOriginalSourceOffsets(t *testing.T) {
	t.Parallel()
	source := "Sub Probe()\r\n    Select Case value\r\n    Case Else\r\n    Case 1\r\n    End Select\r\nEnd Sub\r\n"
	doc, err := vbaast.ParseDocument("Main.bas", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	err = doc.Read(func(view vbaast.ParsedView) error {
		issues := selectCaseSyntaxIssues(string(view.Source), view.Root)
		if len(issues) != 1 {
			t.Fatalf("issues = %#v, want one case-after-else issue", issues)
		}
		issue := issues[0]
		if issue.Kind != "case_after_else" {
			t.Fatalf("kind = %q, want case_after_else", issue.Kind)
		}
		wantStart := strings.Index(source, "Case 1")
		if issue.Range.StartByte != wantStart || source[issue.Range.StartByte:issue.Range.EndByte] != "Case 1" {
			t.Fatalf("range = %#v, text %q; want source offset %d", issue.Range, source[issue.Range.StartByte:issue.Range.EndByte], wantStart)
		}
		return nil
	})
	if err != nil {
		t.Fatal(err)
	}
}

func TestLinterPublishesSelectCaseSyntaxDiagnostic(t *testing.T) {
	source := []byte("Sub Probe()\n    Select Case value\n    Case Else\n        Debug.Print \"default\"\n    Case Else\n        Debug.Print \"duplicate\"\n    End Select\nEnd Sub\n")
	issues, err := (Linter{Config: config.Default()}).LintSource("Main.bas", source)
	if err != nil {
		t.Fatal(err)
	}
	got := issuesByCode(issues, selectCaseDiagnosticCode)
	if len(got) != 1 {
		t.Fatalf("VB063 = %+v, want one diagnostic; issues = %+v", got, issues)
	}
	if got[0].Kind != "duplicate_case_else" || got[0].Line != 5 || got[0].EndLine != 5 {
		t.Fatalf("VB063 = %+v, want duplicate Case Else range", got[0])
	}
	if recovery := issuesByCode(issues, "VB014"); len(recovery) != 0 {
		t.Fatalf("specific Select/Case syntax should not add VB014: %+v", recovery)
	}
}

type selectCaseExpectation struct {
	kind string
	line int
}
