package analyze

import (
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/sourceproject"
)

type selectCaseUnreachableExpectation struct {
	kind string
	line int
	item string
}

func selectCaseUnreachableSummaries(findings []Finding) []selectCaseUnreachableExpectation {
	var got []selectCaseUnreachableExpectation
	for _, finding := range findingsByCode(findings, "VBA259") {
		expectation := selectCaseUnreachableExpectation{line: finding.Line}
		if finding.SelectCaseUnreachable != nil {
			expectation.kind = finding.SelectCaseUnreachable.Kind
			expectation.item = finding.SelectCaseUnreachable.Item
		}
		got = append(got, expectation)
	}
	return got
}

func TestAnalyzerSelectCaseUnreachableFindings(t *testing.T) {
	tests := []struct {
		name   string
		source string
		want   []selectCaseUnreachableExpectation
	}{
		{
			name: "duplicate singleton",
			source: `Option Explicit
Public Sub Run(ByVal x As Long)
  Select Case x
    Case 1
      x = 1
    Case 1
      x = 2
  End Select
End Sub
`,
			want: []selectCaseUnreachableExpectation{{kind: "duplicate", line: 6, item: "1"}},
		},
		{
			name: "covered singleton",
			source: `Option Explicit
Public Sub Run(ByVal x As Long)
  Select Case x
    Case 1 To 10
      x = 1
    Case 5
      x = 2
  End Select
End Sub
`,
			want: []selectCaseUnreachableExpectation{{kind: "covered", line: 6, item: "5"}},
		},
		{
			name: "covered range",
			source: `Option Explicit
Public Sub Run(ByVal x As Long)
  Select Case x
    Case 1 To 10
      x = 1
    Case 3 To 6
      x = 2
  End Select
End Sub
`,
			want: []selectCaseUnreachableExpectation{{kind: "covered", line: 6, item: "3 To 6"}},
		},
		{
			name: "empty range",
			source: `Option Explicit
Public Sub Run(ByVal x As Long)
  Select Case x
    Case 10 To 5
      x = 1
    Case 1
      x = 2
  End Select
End Sub
`,
			want: []selectCaseUnreachableExpectation{{kind: "empty_range", line: 4, item: "10 To 5"}},
		},
		{
			name: "impossible is comparison on byte",
			source: `Option Explicit
Public Sub Run(ByVal x As Byte)
  Select Case x
    Case Is > 300
      x = 1
    Case 0
      x = 2
  End Select
End Sub
`,
			want: []selectCaseUnreachableExpectation{{kind: "impossible", line: 4, item: "Is > 300"}},
		},
		{
			name: "else covered by boolean domain",
			source: `Option Explicit
Public Sub Run(ByVal flag As Boolean)
  Select Case flag
    Case True
      flag = False
    Case False
      flag = True
    Case Else
      flag = True
  End Select
End Sub
`,
			want: []selectCaseUnreachableExpectation{{kind: "else_covered", line: 8}},
		},
		{
			name: "else covered by byte range",
			source: `Option Explicit
Public Sub Run(ByVal x As Byte)
  Select Case x
    Case 0 To 255
      x = 1
    Case Else
      x = 2
  End Select
End Sub
`,
			want: []selectCaseUnreachableExpectation{{kind: "else_covered", line: 6}},
		},
		{
			name: "multiple items in one clause",
			source: `Option Explicit
Public Sub Run(ByVal x As Long)
  Select Case x
    Case 1, 2 To 5
      x = 1
    Case 3, 9
      x = 2
  End Select
End Sub
`,
			want: []selectCaseUnreachableExpectation{{kind: "covered", line: 6, item: "3"}},
		},
		{
			name: "nested select tracked independently",
			source: `Option Explicit
Public Sub Run(ByVal x As Long, ByVal y As Long)
  Select Case x
    Case 1
      Select Case y
        Case 1
          y = 1
        Case 1
          y = 2
      End Select
    Case 1
      x = 2
  End Select
End Sub
`,
			want: []selectCaseUnreachableExpectation{
				{kind: "duplicate", line: 8, item: "1"},
				{kind: "duplicate", line: 11, item: "1"},
			},
		},
		{
			name: "option compare binary keeps case-distinct strings reachable",
			source: `Option Explicit
Option Compare Binary
Public Sub Run(ByVal s As String)
  Select Case s
    Case "ABC"
      s = "a"
    Case "abc"
      s = "b"
    Case "abc"
      s = "c"
  End Select
End Sub
`,
			want: []selectCaseUnreachableExpectation{{kind: "duplicate", line: 9, item: `"abc"`}},
		},
		{
			name: "option compare text folds ascii case",
			source: `Option Explicit
Option Compare Text
Public Sub Run(ByVal s As String)
  Select Case s
    Case "ABC"
      s = "a"
    Case "abc"
      s = "b"
  End Select
End Sub
`,
			want: []selectCaseUnreachableExpectation{{kind: "duplicate", line: 7, item: `"abc"`}},
		},
		{
			name: "option compare database fails open",
			source: `Option Explicit
Option Compare Database
Public Sub Run(ByVal s As String)
  Select Case s
    Case "abc"
      s = "a"
    Case "abc"
      s = "b"
  End Select
End Sub
`,
		},
		{
			name: "byte domain impossibility",
			source: `Option Explicit
Public Sub Run(ByVal x As Byte)
  Select Case x
    Case 300
      x = 1
    Case -1
      x = 2
    Case 0 To 255
      x = 3
    Case Else
      x = 4
  End Select
End Sub
`,
			want: []selectCaseUnreachableExpectation{
				{kind: "impossible", line: 4, item: "300"},
				{kind: "impossible", line: 6, item: "-1"},
				{kind: "else_covered", line: 10},
			},
		},
		{
			name: "boolean domain impossibility",
			source: `Option Explicit
Public Sub Run(ByVal flag As Boolean)
  Select Case flag
    Case 2
      flag = True
    Case True
      flag = False
  End Select
End Sub
`,
			want: []selectCaseUnreachableExpectation{{kind: "impossible", line: 4, item: "2"}},
		},
		{
			name: "select case true singleton",
			source: `Option Explicit
Public Sub Run()
  Select Case True
    Case -1
      Debug.Print "yes"
    Case False
      Debug.Print "no"
    Case Else
      Debug.Print "other"
  End Select
End Sub
`,
			want: []selectCaseUnreachableExpectation{
				{kind: "impossible", line: 6, item: "False"},
				{kind: "else_covered", line: 8},
			},
		},
		{
			name: "variant selector still reports exact duplicates",
			source: `Option Explicit
Public Sub Run(ByVal x As Variant)
  Select Case x
    Case 1
      x = 1
    Case 1
      x = 2
    Case Else
      x = 3
  End Select
End Sub
`,
			want: []selectCaseUnreachableExpectation{{kind: "duplicate", line: 6, item: "1"}},
		},
		{
			name: "cross-type item fails open",
			source: `Option Explicit
Public Sub Run(ByVal s As String)
  Select Case s
    Case 5
      s = "a"
    Case 5
      s = "b"
  End Select
End Sub
`,
		},
		{
			name: "opaque date selector fails open",
			source: `Option Explicit
Public Sub Run(ByVal d As Date)
  Select Case d
    Case 1
      d = Date
    Case 1
      d = Date
  End Select
End Sub
`,
		},
		{
			name: "unknown case operand fails open",
			source: `Option Explicit
Public Sub Run(ByVal x As Long)
  Select Case x
    Case OtherValue(x)
      x = 1
    Case OtherValue(x)
      x = 2
  End Select
End Sub
`,
		},
		{
			name: "module constant operand participates in coverage",
			source: `Option Explicit
Private Const Limit As Long = 5
Public Sub Run(ByVal x As Long)
  Select Case x
    Case Limit
      x = 1
    Case 5
      x = 2
    Case Else
      x = 3
  End Select
End Sub
`,
			want: []selectCaseUnreachableExpectation{{kind: "duplicate", line: 7, item: "5"}},
		},
		{
			name: "constant selector reports impossible and covered else",
			source: `Option Explicit
Private Const Mode As Long = 7
Public Sub Run()
  Select Case Mode
    Case 7
      Debug.Print "hit"
    Case 9
      Debug.Print "miss"
    Case Else
      Debug.Print "other"
  End Select
End Sub
`,
			want: []selectCaseUnreachableExpectation{
				{kind: "impossible", line: 7, item: "9"},
				{kind: "else_covered", line: 9},
			},
		},
		{
			name: "clauses after case else belong to VB063",
			source: `Option Explicit
Public Sub Run(ByVal x As Long)
  Select Case x
    Case Else
      x = 1
    Case 1
      x = 2
  End Select
End Sub
`,
		},
		{
			name: "conditional compilation fails open",
			source: `Option Explicit
Public Sub Run(ByVal x As Long)
  Select Case x
    Case 1
      x = 1
#If Mac Then
    Case 1
#End If
      x = 2
  End Select
End Sub
`,
		},
	}

	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			dir := t.TempDir()
			writeModule(t, dir, "Main.bas", test.source)
			cfg := config.Default()
			cfg.Analyze.DetectUnreachableSelectCase = true
			findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
			if err != nil {
				t.Fatal(err)
			}
			got := selectCaseUnreachableSummaries(findings)
			if !reflect.DeepEqual(got, test.want) {
				t.Fatalf("VBA259 findings = %+v, want %+v (all findings: %+v)", got, test.want, findings)
			}
		})
	}
}

func TestAnalyzerSelectCaseUnreachableDefaultOff(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal x As Long)
  Select Case x
    Case 1
      x = 1
    Case 1
      x = 2
  End Select
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA259"); len(got) != 0 {
		t.Fatalf("VBA259 findings = %+v, want none while the rule is disabled by default", got)
	}
}

func TestAnalyzerSelectCaseUnreachableInlineSuppression(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal x As Long)
  Select Case x
    Case 1
      x = 1
    ' xlflow:disable-next-line VBA259
    Case 1
      x = 2
    Case 1
      x = 3
  End Select
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectUnreachableSelectCase = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := selectCaseUnreachableSummaries(findings)
	want := []selectCaseUnreachableExpectation{{kind: "duplicate", line: 9, item: "1"}}
	if !reflect.DeepEqual(got, want) {
		t.Fatalf("VBA259 findings = %+v, want only the unsuppressed duplicate on line 9", got)
	}
}

func TestAnalyzerSelectCaseUnreachableFilesystemInMemoryParity(t *testing.T) {
	source := `Option Explicit
Public Sub Run(ByVal x As Long)
  Select Case x
    Case 1 To 10
      x = 1
    Case 5
      x = 2
    Case Else
      x = 3
  End Select
End Sub
`
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", source)
	cfg := config.Default()
	cfg.Analyze.DetectUnreachableSelectCase = true

	filesystemFindings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatalf("filesystem analysis: %v", err)
	}
	inMemoryResult, err := (Analyzer{RootDir: t.TempDir(), Config: cfg}).AnalyzeProject(t.Context(), sourceproject.SourceProject{
		Files: []sourceproject.SourceFile{{
			Path:       "src/modules/Main.bas",
			Source:     []byte(source),
			ModuleKind: sourceproject.ModuleKindStandard,
		}},
	})
	if err != nil {
		t.Fatalf("in-memory analysis: %v", err)
	}

	filesystemGot := selectCaseUnreachableSummaries(filesystemFindings)
	inMemoryGot := selectCaseUnreachableSummaries(inMemoryResult.Findings)
	if !reflect.DeepEqual(filesystemGot, inMemoryGot) {
		t.Fatalf("VBA259 findings differ: filesystem %+v, in-memory %+v", filesystemGot, inMemoryGot)
	}
	if len(filesystemGot) != 1 || filesystemGot[0].kind != "covered" || filesystemGot[0].line != 6 {
		t.Fatalf("VBA259 findings = %+v, want one covered finding on line 6", filesystemGot)
	}
}

func TestAnalyzerSelectCaseUnreachableFindingCarriesContext(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal x As Long)
  Select Case x
    Case 1
      x = 1
    Case 1
      x = 2
  End Select
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectUnreachableSelectCase = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA259")
	if len(got) != 1 {
		t.Fatalf("VBA259 findings = %+v, want one duplicate", got)
	}
	finding := got[0]
	if finding.SelectCaseUnreachable == nil {
		t.Fatalf("finding missing SelectCaseUnreachable context: %+v", finding)
	}
	if finding.SelectCaseUnreachable.Kind != "duplicate" || finding.SelectCaseUnreachable.Item != "1" || finding.SelectCaseUnreachable.CoveredByLine != 4 {
		t.Fatalf("unexpected context: %+v", finding.SelectCaseUnreachable)
	}
	if finding.Severity != "warning" || !strings.Contains(finding.Message, `"1"`) {
		t.Fatalf("unexpected finding payload: %+v", finding)
	}
}
