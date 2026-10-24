package intel

import (
	"os"
	"path/filepath"
	"reflect"
	"slices"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/lint"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
)

func TestProcedureNameConstantQuickFixMatchesPublishedDiagnostic(t *testing.T) {
	for _, literal := range []string{`"OldName"`, `"日本語😀""旧名"`} {
		for _, eol := range []string{"\n", "\r\n"} {
			t.Run(literal+"/"+strings.ReplaceAll(eol, "\n", "LF"), func(t *testing.T) {
				analyzer := newTestAnalyzer(t)
				analyzer.Config.Lint.ProcedureNameConstant.Enabled = true
				analyzer.Config.Lint.ProcedureNameConstant.ConstantName = "PROCEDURE_NAME"
				source := strings.ReplaceAll("Option Explicit\nPublic Sub Foo()\n    Const PROCEDURE_NAME As String = "+literal+"\nEnd Sub\n", "\n", eol)
				doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: source}
				diagnostics := analyzer.Diagnostics(doc)
				index := slices.IndexFunc(diagnostics, func(d Diagnostic) bool { return d.Code == "VB044" })
				if index < 0 {
					t.Fatalf("VB044 diagnostic missing: %+v", diagnostics)
					return
				}
				published := diagnostics[index]
				start := strings.Index(source, literal)
				wantRange := Range{Start: positionForByteOffset(source, start), End: positionForByteOffset(source, start+len(literal))}
				actions, err := analyzer.CodeActionsContext(t.Context(), doc, wantRange, []string{"quickfix"})
				if err != nil || len(actions) != 1 {
					t.Fatalf("actions = %+v, err = %v", actions, err)
					return
				}
				if !reflect.DeepEqual(actions[0].Diagnostics, []Diagnostic{published}) {
					t.Fatalf("associated diagnostics = %+v, want published %+v", actions[0].Diagnostics, published)
				}
				if len(actions[0].Edits) != 1 || actions[0].Edits[0].Range != wantRange {
					t.Fatalf("edits = %+v, want full literal range %+v", actions[0].Edits, wantRange)
				}
			})
		}
	}
}

func TestRedundantStepOneQuickFixPreservesSource(t *testing.T) {
	for _, tc := range []struct {
		name, header, want string
	}{
		{"ordinary", "For i = 1 To 3 Step 1", "For i = 1 To 3"},
		{"tabs and case", "For i = 1 To 3\tStEp\t1", "For i = 1 To 3"},
		{"comment", "For i = 1 To 3 Step 1 ' 日本語 😀", "For i = 1 To 3 ' 日本語 😀"},
		{"colon and unicode", "Debug.Print \"日本語😀\": For i = 1 To 3 Step 1", "Debug.Print \"日本語😀\": For i = 1 To 3"},
		{"continued bound", "For i = 1 To _\n    3 Step 1", "For i = 1 To _\n    3"},
		{"conditional", "#If VBA7 Then\nFor i = 1 To 3 Step 1", "#If VBA7 Then\nFor i = 1 To 3"},
	} {
		for _, eol := range []string{"\n", "\r\n"} {
			t.Run(tc.name+"/"+strings.ReplaceAll(eol, "\n", "LF"), func(t *testing.T) {
				analyzer := newTestAnalyzer(t)
				analyzer.Config.Lint.DetectRedundantStepOne = true
				suffix := "\n    Debug.Print i\nNext i\n"
				if tc.name == "conditional" {
					suffix += "#End If\n"
				}
				source := strings.ReplaceAll("Option Explicit\nSub Work()\nDim i As Long\n"+tc.header+suffix+"End Sub\n", "\n", eol)
				want := strings.ReplaceAll("Option Explicit\nSub Work()\nDim i As Long\n"+tc.want+suffix+"End Sub\n", "\n", eol)
				doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: source, Version: 3}
				snapshot := NewAnalysisSnapshot(doc)
				defer snapshot.Retire()
				doc = snapshot.Document()
				start := strings.Index(strings.ToLower(source), "step")
				selection := positionForDocumentByteOffset(doc, start)
				actions, err := analyzer.CodeActionsContext(t.Context(), doc, Range{Start: selection, End: selection}, []string{"quickfix"})
				if err != nil {
					t.Fatal(err)
				}
				if len(actions) != 1 || len(actions[0].Diagnostics) != 1 || actions[0].Diagnostics[0].Code != "VB084" {
					t.Fatalf("actions = %+v", actions)
					return
				}
				got := applyQuickFixForTest(t, doc, actions[0])
				if got != want {
					t.Fatalf("corrected source = %q, want %q", got, want)
				}
				if stats := snapshot.ProcedureArtifactStats(); stats.IRBuild != 0 || stats.CFGBuild != 0 {
					t.Fatalf("correction discovery built procedure artifacts: %+v", stats)
				}
				parsed, err := vbaast.ParseDocument(doc.Path, []byte(got))
				if err != nil {
					t.Fatal(err)
				}
				defer parsed.Close()
				issues, err := (lint.Linter{RootDir: analyzer.RootDir, Config: analyzer.Config}).LintParsed(parsed)
				if err != nil {
					t.Fatal(err)
				}
				for _, issue := range issues {
					if issue.Code == "VB084" || issue.Code == "VB014" {
						t.Fatalf("corrected source has unexpected diagnostic: %+v", issue)
					}
				}
				doc = Document{Path: doc.Path, Source: got}
				actions, err = analyzer.CodeActionsContext(t.Context(), doc, Range{End: positionForByteOffset(got, len(got))}, []string{"quickfix"})
				if err != nil || len(actions) != 0 {
					t.Fatalf("second correction = %+v, %v", actions, err)
				}
			})
		}
	}
}

func TestRedundantStepOneQuickFixWithholdsUnsafeOrInvisibleCorrections(t *testing.T) {
	for _, tc := range []struct {
		name, header, before string
		disabled, unrelated  bool
	}{
		{name: "disabled", header: "For i = 1 To 3 Step 1", disabled: true},
		{name: "other step", header: "For i = 1 To 3 Step 2"},
		{name: "expression", header: "For i = 1 To 3 Step 1 + 0"},
		{name: "parenthesized", header: "For i = 1 To 3 Step (1)"},
		{name: "unary", header: "For i = 1 To 3 Step +1"},
		{name: "continued keyword", header: "For i = 1 To 3 _\n    Step 1"},
		{name: "continued literal", header: "For i = 1 To 3 Step _\n    1"},
		{name: "malformed", header: "For i = To 3 Step 1"},
		{name: "disable line", header: "For i = 1 To 3 Step 1 ' xlflow:disable-line VB084"},
		{name: "disable next line", before: "' xlflow:disable-next-line VB084\n", header: "For i = 1 To 3 Step 1"},
		{name: "unrelated selection", header: "For i = 1 To 3 Step 1", unrelated: true},
	} {
		t.Run(tc.name, func(t *testing.T) {
			analyzer := newTestAnalyzer(t)
			analyzer.Config.Lint.DetectRedundantStepOne = !tc.disabled
			if tc.disabled {
				analyzer.Config.Lint.RequireOptionExplicit = false
				analyzer.DB = nil
			}
			source := "Option Explicit\nSub Work()\nDim i As Long\n" + tc.before + tc.header + "\nDebug.Print i\nNext i\nEnd Sub\n"
			doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: source}
			snapshot := NewAnalysisSnapshot(doc)
			defer snapshot.Retire()
			doc = snapshot.Document()
			selection := Range{End: positionForByteOffset(source, len(source))}
			if tc.unrelated {
				selection = Range{}
			}
			actions, err := analyzer.CodeActionsContext(t.Context(), doc, selection, []string{"quickfix"})
			if err != nil || len(actions) != 0 {
				t.Fatalf("actions = %+v, err = %v", actions, err)
			}
			if tc.disabled && snapshot.ParseCount() != 0 {
				t.Fatal("disabled corrections parsed the document")
			}
		})
	}
}

func TestRedundantStepOneQuickFixMatchesVBEFixtures(t *testing.T) {
	root := filepath.Join("..", "..", "..", "testdata", "vbe-oracle", "cases")
	before, err := os.ReadFile(filepath.Join(root, "quickfix-step-one-before", "Main.bas"))
	if err != nil {
		t.Fatal(err)
	}
	after, err := os.ReadFile(filepath.Join(root, "quickfix-step-one-after", "Main.bas"))
	if err != nil {
		t.Fatal(err)
	}
	analyzer := newTestAnalyzer(t)
	analyzer.Config.Lint.DetectRedundantStepOne = true
	doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: string(before)}
	actions, err := analyzer.CodeActionsContext(t.Context(), doc, Range{End: positionForByteOffset(doc.Source, len(doc.Source))}, []string{"quickfix"})
	if err != nil || len(actions) != 2 {
		t.Fatalf("fixture actions = %+v, %v", actions, err)
	}
	var combined CodeAction
	for _, action := range actions {
		combined.Edits = append(combined.Edits, action.Edits...)
	}
	if got := applyQuickFixForTest(t, doc, combined); got != string(after) {
		t.Fatalf("correction differs from VBE-confirmed source: %q", got)
	}
}

func applyQuickFixForTest(t *testing.T, doc Document, action CodeAction) string {
	t.Helper()
	edits := slices.Clone(action.Edits)
	slices.SortFunc(edits, func(a, b CodeActionEdit) int {
		return byteOffsetForDocumentPosition(doc, b.Range.Start) - byteOffsetForDocumentPosition(doc, a.Range.Start)
	})
	result := doc.Source
	for _, edit := range edits {
		start := byteOffsetForDocumentPosition(doc, edit.Range.Start)
		end := byteOffsetForDocumentPosition(doc, edit.Range.End)
		result = result[:start] + edit.NewText + result[end:]
	}
	return result
}
