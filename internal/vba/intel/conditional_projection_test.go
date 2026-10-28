package intel

import (
	"context"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/lint"
	"github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/conditional"
)

func TestDiagnosticsProjectionMatchesBatchLintForConditionalProse(t *testing.T) {
	tests := []struct {
		name      string
		constants map[string]string
		wantVB014 bool
	}{
		{
			name: "undefined custom constant is false",
		},
		{
			name:      "explicit false selects else",
			constants: map[string]string{"EnableBranch": "False"},
		},
		{
			name:      "explicit true selects prose",
			constants: map[string]string{"EnableBranch": "True"},
			wantVB014: true,
		},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			cfg := config.Default()
			cfg.VBA.ConditionalConstants = test.constants
			source := "#If EnableBranch Then\r\n" +
				"This is plain prose, not VBA.\r\n" +
				"## Heading\r\n" +
				"#Else\r\n" +
				"Public Sub Selected()\r\nEnd Sub\r\n" +
				"#End If\r\n" +
				"Public Function Answer() As Long\r\n    Answer = 42\r\nEnd Function\r\n"
			path := "Main.bas"

			batchIssues, err := (lint.Linter{Config: cfg}).LintSourceContext(context.Background(), path, []byte(source))
			if err != nil {
				t.Fatal(err)
			}
			batchHasVB014 := containsIssueCode(batchIssues, "VB014")
			if batchHasVB014 != test.wantVB014 {
				t.Fatalf("batch VB014 = %v, want %v; issues=%+v", batchHasVB014, test.wantVB014, batchIssues)
			}

			diagnostics := (Analyzer{Config: cfg}).Diagnostics(Document{Path: path, Source: source})
			lspHasVB014 := containsDiagnosticCode(diagnostics, "VB014")
			if lspHasVB014 != batchHasVB014 {
				t.Fatalf("LSP VB014 = %v, batch VB014 = %v; diagnostics=%+v issues=%+v", lspHasVB014, batchHasVB014, diagnostics, batchIssues)
			}
		})
	}
}

func TestConditionalProjectionPreservesRawSourceAndUTF16CRLFRanges(t *testing.T) {
	source := "#If NotDefinedConst Then\r\n" +
		"' 日本語 😀\r\n" +
		"This is inactive prose.\r\n" +
		"#End If\r\n" +
		"Public Sub Active()\r\n" +
		"    Debug.Print \"東京😀\": TypeOf value Collection Extra: Debug.Print (1\r\n" +
		"End Sub\r\n"
	doc := Document{Path: "Main.bas", Source: source}.WithConditionalProjection(nil)
	snapshot := NewAnalysisSnapshot(doc)
	defer snapshot.Retire()
	if snapshot.Source() != source || snapshot.Document().Source != source {
		t.Fatal("snapshot did not retain the original source")
	}
	projected := snapshot.AnalysisSource()
	if len(projected) != len(source) || strings.Count(projected, "\r\n") != strings.Count(source, "\r\n") {
		t.Fatalf("projection changed source shape: original=%d/%d CRLF projected=%d/%d CRLF", len(source), strings.Count(source, "\r\n"), len(projected), strings.Count(projected, "\r\n"))
	}
	if strings.Contains(projected, "This is inactive prose") {
		t.Fatal("inactive prose remained in the projected source")
	}

	diagnostics := (Analyzer{Config: config.Default()}).Diagnostics(snapshot.Document())
	var got *Diagnostic
	for i := range diagnostics {
		if diagnostics[i].Code == "VB014" {
			got = &diagnostics[i]
			break
		}
	}
	if got == nil {
		t.Fatalf("diagnostics = %+v, want VB014 for active malformed statement", diagnostics)
	}
	batchIssues, err := (lint.Linter{Config: config.Default()}).LintSourceContext(context.Background(), doc.Path, []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	var expected Range
	for _, issue := range batchIssues {
		if issue.Code == "VB014" {
			expected = issueRange(source, issue.Line, issue.Column)
			break
		}
	}
	if got.Range != expected {
		t.Fatalf("LSP VB014 range = %+v, batch-mapped range = %+v", got.Range, expected)
	}
	if got.Range.Start.Line != 5 {
		t.Fatalf("VB014 line = %d, want 5 (CRLF lines after masked Unicode)", got.Range.Start.Line)
	}
	invalidToken := strings.Index(source, "TypeOf")
	lineStart := strings.LastIndex(source[:invalidToken], "\r\n") + 2
	wantCharacter := utf16Len(source[lineStart:invalidToken])
	if got.Range.Start.Character <= utf16Len(source[lineStart:strings.Index(source[lineStart:], "😀")+lineStart+len("😀")]) {
		t.Fatalf("VB014 range character = %d, want a location after the astral Unicode text", got.Range.Start.Character)
	}
	if got.Range.Start.Character != wantCharacter {
		t.Fatalf("VB014 UTF-16 character = %d, want %d after Unicode text", got.Range.Start.Character, wantCharacter)
	}
}

func TestConditionalProjectionRecognizesCROnlyDirectives(t *testing.T) {
	source := "Option Explicit\r#If False Then\rInactive prose\r#End If\r"
	doc := Document{Path: "Main.bas", Source: source}.WithConditionalProjection(nil)
	if strings.Contains(doc.AnalysisSource(), "Inactive prose") {
		t.Fatal("CR-only source bypassed conditional projection")
	}
	if doc.Source != source {
		t.Fatal("authored source changed")
	}
}

func TestConditionalProjectionSnapshotIdentityIncludesConstants(t *testing.T) {
	doc := Document{Path: "Main.bas", Source: "#If Feature Then\nThis is prose.\n#Else\nSub Active()\nEnd Sub\n#End If\n"}.
		WithConditionalProjection(map[string]string{"Feature": "True"})
	snapshot := NewAnalysisSnapshot(doc)
	defer snapshot.Retire()
	if !snapshot.Matches(doc) {
		t.Fatal("snapshot did not match its captured conditional environment")
	}
	changed := doc.WithConditionalProjection(map[string]string{"Feature": "False"})
	if snapshot.Matches(changed) {
		t.Fatal("snapshot matched a different conditional environment")
	}
	if got := changed.AnalysisSource(); strings.Contains(got, "This is prose.") {
		t.Fatal("explicit false did not mask the active prose branch")
	}
}

func TestConditionalProjectionReusesSeededSnapshotAndDocumentIndex(t *testing.T) {
	constants := map[string]string{"Feature": "True"}
	rawSource := "#If Feature Then\n" +
		"Public Sub Active()\nEnd Sub\n" +
		"#Else\n" +
		"Public Sub Inactive()\nEnd Sub\n" +
		"#End If\n"
	analysisSource := conditional.Project([]byte(rawSource), constants)
	parsed, err := ast.ParseDocument("Main.bas", analysisSource)
	if err != nil {
		t.Fatal(err)
	}
	doc := Document{
		Path:                 "Main.bas",
		Source:               string(analysisSource),
		ProjectConditionals:  true,
		ConditionalConstants: constants,
	}
	snapshot := NewAnalysisSnapshotWithParsedDocument(doc, parsed)
	defer snapshot.RetireAndWait()
	doc = snapshot.Document()
	cfg := config.Default()
	cfg.VBA.ConditionalConstants = constants
	analyzer := Analyzer{Config: cfg}

	for _, diagnostics := range []func(Document){
		func(doc Document) { analyzer.CompileEquivalentDiagnosticsContext(context.Background(), doc) },
		func(doc Document) {
			analyzer.CompileEquivalentDiagnosticsContextWithByRefDiagnostics(context.Background(), doc, nil)
		},
		func(doc Document) { analyzer.Diagnostics(doc) },
	} {
		diagnostics(doc)
		projected := analyzer.projectConditionalDocument(doc)
		if projected.Snapshot != snapshot || !snapshot.Matches(projected) {
			t.Fatal("diagnostic entrypoint discarded the seeded projection snapshot")
		}
	}

	var firstIndex *documentIndex
	for range 3 {
		index, ok := analyzer.documentIndexFor(doc)
		if !ok || index == nil {
			t.Fatal("document index was unavailable")
		}
		if firstIndex == nil {
			firstIndex = index
		} else if index != firstIndex {
			t.Fatal("document index was rebuilt for the same projection snapshot")
		}
	}
	if got := snapshot.ParseCount(); got != 1 {
		t.Fatalf("snapshot parse count = %d, want one seeded parse", got)
	}
}

func TestConditionalProjectionDocumentIndexIgnoresInactiveAssignments(t *testing.T) {
	constants := map[string]string{"Feature": "True"}
	source := "#If Feature Then\n" +
		"Public Sub Run()\n" +
		"    Dim item As Object\n" +
		"    Set item = New Collection\n" +
		"End Sub\n" +
		"#Else\n" +
		"Public Sub Run()\n" +
		"    Dim item As Object\n" +
		"    Set item = New Worksheet\n" +
		"End Sub\n" +
		"#End If\n"
	doc := Document{Path: "Main.bas", Source: source}.WithConditionalProjection(constants)
	snapshot := NewAnalysisSnapshot(doc)
	defer snapshot.RetireAndWait()
	doc = snapshot.Document()
	index, ok := (Analyzer{Config: config.Default()}).documentIndexFor(doc)
	if !ok || index == nil {
		t.Fatal("document index was unavailable")
	}
	index.initAssignments()
	assignments := index.assignmentsByName[indexName("item")]
	if len(assignments) != 1 || !strings.Contains(assignments[0].expression, "New Collection") {
		t.Fatalf("indexed assignments = %+v, want only the active Collection assignment", assignments)
	}
}

func TestConditionalProjectionLeavesUnconditionalSnapshotReusable(t *testing.T) {
	doc := Document{Path: "Main.bas", Source: "Public Sub Run()\nEnd Sub\n"}
	snapshot := NewAnalysisSnapshot(doc)
	defer snapshot.RetireAndWait()
	doc = snapshot.Document()
	projected := (Analyzer{Config: config.Default()}).projectConditionalDocument(doc)
	if projected.Snapshot != snapshot || !snapshot.Matches(projected) {
		t.Fatal("unconditional source invalidated its existing snapshot")
	}
}

func TestAnalysisSnapshotLeavesStructuralSourceUnprojectedByDefault(t *testing.T) {
	source := "#If NotDefinedConst Then\nThis is prose.\n#End If\n"
	snapshot := NewAnalysisSnapshot(Document{Path: "Main.bas", Source: source})
	defer snapshot.Retire()
	if snapshot.AnalysisSource() != source {
		t.Fatal("zero-value structural snapshot projected conditional source")
	}
}

func containsIssueCode(issues []lint.Issue, code string) bool {
	for _, issue := range issues {
		if issue.Code == code {
			return true
		}
	}
	return false
}

func containsDiagnosticCode(diagnostics []Diagnostic, code string) bool {
	for _, diagnostic := range diagnostics {
		if diagnostic.Code == code {
			return true
		}
	}
	return false
}
