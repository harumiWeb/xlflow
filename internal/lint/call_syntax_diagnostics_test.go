package lint

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestCallSyntaxDiagnostics(t *testing.T) {
	tests := []struct {
		name   string
		source string
		want   map[string]int
	}{
		{
			name: "explicit call requires parentheses",
			source: `Sub Probe()
    Call WritePair 1, 2
End Sub
Private Sub WritePair(ByVal first As Long, ByVal second As Long)
End Sub
`,
			want: map[string]int{"explicit_call_requires_parentheses": 1},
		},
		{
			name: "standalone empty parentheses",
			source: `Sub Probe()
    WriteValue()
End Sub
Private Sub WriteValue()
End Sub
`,
			want: map[string]int{"standalone_empty_parentheses": 1},
		},
		{
			name: "standalone multi argument parentheses",
			source: `Sub Probe()
    WritePair (1, 2)
End Sub
Private Sub WritePair(ByVal first As Long, ByVal second As Long)
End Sub
`,
			want: map[string]int{"standalone_multi_parenthesized": 1},
		},
		{
			name: "standalone omitted multi argument parentheses",
			source: `Sub Probe()
    WritePair (, 2)
End Sub
Private Sub WritePair(ByRef first As Long, ByRef second As Long)
End Sub
`,
			want: map[string]int{"standalone_multi_parenthesized": 1},
		},
		{
			name: "function expression without parentheses",
			source: `Sub Probe()
    Dim value As Long
    value = AddOne 1
End Sub
Private Function AddOne(ByVal value As Long) As Long
    AddOne = value + 1
End Function
`,
			want: map[string]int{"function_expression_requires_parentheses": 1},
		},
		{
			name: "elseif function expression without parentheses",
			source: `Sub Probe()
    If True Then
    ElseIf AddOne 1 Then
    End If
End Sub
Private Function AddOne(ByVal value As Long) As Boolean
    AddOne = value > 0
End Function
`,
			want: map[string]int{"function_expression_requires_parentheses": 1},
		},
		{
			name: "invalid explicit target",
			source: `Sub Probe()
    Call (1 + 2)
End Sub
`,
			want: map[string]int{"invalid_call_target": 1},
		},
		{
			name: "invalid explicit target expression",
			source: `Sub Probe()
    Call Foo + Bar
End Sub
`,
			want: map[string]int{"invalid_call_target": 1},
		},
		{
			name: "invalid explicit target after arguments",
			source: `Sub Probe()
    Call Foo(1) + Bar
End Sub
`,
			want: map[string]int{"invalid_call_target": 1},
		},
		{
			name:   "invalid explicit target after unrelated continuation",
			source: "Sub Probe()\n    Dim value As Long\n    value = 1 + _\n        2\n    Call Foo(1) + Bar\nEnd Sub\n",
			want:   map[string]int{"invalid_call_target": 1},
		},
		{
			name: "invalid qualified explicit target",
			source: `Sub Probe()
    Call Foo..Bar
End Sub
`,
			want: map[string]int{"invalid_call_target": 1},
		},
		{
			name: "legal parenthesized argument idiom",
			source: `Sub Probe()
    WritePair (1), (2)
End Sub
Private Sub WritePair(ByRef first As Long, ByRef second As Long)
End Sub
`,
			want: map[string]int{},
		},
		{
			name: "legal parenthesized ByVal expressions",
			source: `Sub Probe()
    WritePair (1 + 1), (2 + 2)
End Sub
Private Sub WritePair(ByVal first As Long, ByVal second As Long)
End Sub
`,
			want: map[string]int{},
		},
		{
			name: "legal nested function expression",
			source: `Sub Probe()
    Dim value As Long
    value = AddOne(AddOne(1))
End Sub
Private Function AddOne(ByVal value As Long) As Long
    AddOne = value + 1
End Function
`,
			want: map[string]int{},
		},
		{
			name: "legal default member call",
			source: `Sub Probe()
    Dim value As Long
    value = items(index).Value
End Sub
`,
			want: map[string]int{},
		},
		{
			name: "legal explicit call",
			source: `Sub Probe()
    Call WritePair(1, 2)
End Sub
Private Sub WritePair(ByVal first As Long, ByVal second As Long)
End Sub
`,
			want: map[string]int{},
		},
		{
			name: "legal explicit zero argument call",
			source: `Sub Probe()
    Call WriteValue
End Sub
Private Sub WriteValue()
End Sub
`,
			want: map[string]int{},
		},
		{
			name: "legal explicit zero argument call with comment",
			source: `Sub Probe()
    Call WriteValue ' keep the explicit zero-argument form
End Sub
Private Sub WriteValue()
End Sub
`,
			want: map[string]int{},
		},
		{
			name: "explicit call with numeric line label",
			source: `Sub Probe()
10 Call WritePair(1, 2)
End Sub
Private Sub WritePair(ByVal first As Long, ByVal second As Long)
End Sub
`,
			want: map[string]int{},
		},
		{
			name: "explicit call requiring parentheses with numeric line label",
			source: `Sub Probe()
10 Call WritePair 1, 2
End Sub
Private Sub WritePair(ByVal first As Long, ByVal second As Long)
End Sub
`,
			want: map[string]int{"explicit_call_requires_parentheses": 1},
		},
		{
			name:   "explicit call continuation with CRLF",
			source: "Sub Probe()\r\n    Call WritePair(1, _\r\n        2)\r\nEnd Sub\r\nPrivate Sub WritePair(ByVal first As Long, ByVal second As Long)\r\nEnd Sub\r\n",
			want:   map[string]int{},
		},
		{
			name: "legal explicit member call with indexed receiver",
			source: `Sub Probe()
    Call rows(index).Message(first, second)
End Sub
`,
			want: map[string]int{},
		},
		{
			name: "legal explicit implicit member call",
			source: `Sub Probe()
    With record
        Call .Delete
        Call .CopyPicture(1, 2)
    End With
End Sub
`,
			want: map[string]int{},
		},
		{
			name: "legal explicit bang member call without arguments",
			source: `Sub Probe()
    Call record!Message
End Sub
`,
			want: map[string]int{},
		},
		{
			name: "legal explicit bang member call with arguments",
			source: `Sub Probe()
    Call record!Message(1)
End Sub
`,
			want: map[string]int{},
		},
		{
			name: "legal explicit spaced bang member call without arguments",
			source: `Sub Probe()
    Call record ! Message
End Sub
`,
			want: map[string]int{},
		},
		{
			name: "explicit call with split parenthesized arguments",
			source: `Sub Probe()
    Call WritePair (1), (2)
End Sub
Private Sub WritePair(ByVal first As Long, ByVal second As Long)
End Sub
`,
			want: map[string]int{"explicit_call_requires_parentheses": 1},
		},
		{
			name: "legal bare call",
			source: `Sub Probe()
    WritePair 1, 2
End Sub
Private Sub WritePair(ByVal first As Long, ByVal second As Long)
End Sub
`,
			want: map[string]int{},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			issues, err := (Linter{Config: config.Default()}).LintSource("Main.bas", []byte(tt.source))
			if err != nil {
				t.Fatal(err)
			}
			got := map[string]int{}
			for _, issue := range issues {
				if issue.Code == callSyntaxDiagnosticCode {
					got[issue.Kind]++
					if issue.Severity != "error" {
						t.Errorf("VB059 severity = %q, want error", issue.Severity)
					}
				}
			}
			if len(got) != len(tt.want) {
				t.Fatalf("VB059 kinds = %#v, want %#v; all issues = %#v", got, tt.want, issues)
			}
			for kind, want := range tt.want {
				if got[kind] != want {
					t.Errorf("VB059 kind %q count = %d, want %d", kind, got[kind], want)
				}
			}
		})
	}
}

func TestCallSyntaxDiagnosticIsNotConfigurableWithVB022(t *testing.T) {
	source := []byte(`Sub Probe()
    WriteValue()
End Sub
Private Sub WriteValue()
End Sub
`)
	cfg := config.Default()
	cfg.Lint.DetectConfusingCallSyntax = false
	issues, err := (Linter{Config: cfg}).LintSource("Main.bas", source)
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, callSyntaxDiagnosticCode); len(got) != 1 {
		t.Fatalf("VB059 count = %d, want 1; issues = %#v", len(got), issues)
	}
	if got := issuesByCode(issues, "VB022"); len(got) != 0 {
		t.Fatalf("VB022 count = %d, want 0 when disabled", len(got))
	}
	if got := issuesByCode(PushBlockingIssues(issues), callSyntaxDiagnosticCode); len(got) != 1 {
		t.Fatalf("VB059 preflight blockers = %d, want 1", len(got))
	}
}

func TestCallSyntaxDiagnosticDoesNotHideUnrelatedParserRecovery(t *testing.T) {
	source := []byte(`Sub Probe()
    WriteValue()
    If True Then
        WriteValue()
End Sub
Private Sub WriteValue()
End Sub
`)
	issues, err := (Linter{Config: config.Default()}).LintSource("Main.bas", source)
	if err != nil {
		t.Fatal(err)
	}
	if len(issuesByCode(issues, callSyntaxDiagnosticCode)) == 0 {
		t.Fatalf("VB059 missing: %#v", issues)
	}
	if len(issuesByCode(issues, "VB014")) == 0 {
		t.Fatalf("unrelated parser recovery was hidden by VB059: %#v", issues)
	}
}
