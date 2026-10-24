package intel

import (
	"os"
	"path/filepath"
	"reflect"
	"slices"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestOptionExplicitQuickFixPreservesModulePreamble(t *testing.T) {
	for _, tc := range []struct{ name, preamble, body string }{
		{"plain", "", "Public Sub Work()\nEnd Sub\n"},
		{"empty", "", ""},
		{"comment only", "", "' 日本語😀\n"},
		{"procedure docs", "", "''' Work documentation\nPublic Sub Work()\nEnd Sub\n"},
		{"options", "", "Option Compare Text\nOption Base 1\nPublic Sub Work()\nEnd Sub\n"},
		{"standard", "Attribute VB_Name = \"Main\"\nAttribute VB_Description = \"日本語😀\"\n", "' comment\nSub Work()\nEnd Sub\n"},
		{"attribute comment", "Attribute VB_Name = \"Main\" ' module\n", "Sub Work()\nEnd Sub\n"},
		{"class", "VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1\nEND\nAttribute VB_Name = \"Main\"\n", "Sub Work()\nEnd Sub\n"},
		{"form", "VERSION 5.00\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Main\nEnd\nAttribute VB_Name = \"Main\"\n", "Sub Work()\nEnd Sub\n"},
		{"conditional", "", "#If VBA7 Then\nSub Work()\nEnd Sub\n#Else\nSub Work()\nEnd Sub\n#End If\n"},
	} {
		for _, eol := range []string{"\n", "\r\n"} {
			t.Run(tc.name+"/"+strings.ReplaceAll(eol, "\n", "LF"), func(t *testing.T) {
				analyzer := newTestAnalyzer(t)
				source := strings.ReplaceAll(tc.preamble+tc.body, "\n", eol)
				doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: source}
				snapshot := NewAnalysisSnapshot(doc)
				defer snapshot.Retire()
				doc = snapshot.Document()
				actions, err := analyzer.CodeActionsContext(t.Context(), doc, Range{}, []string{"quickfix"})
				if err != nil || len(actions) != 1 {
					t.Fatalf("actions = %+v, err = %v", actions, err)
					return
				}
				if stats := snapshot.ProcedureArtifactStats(); stats.IRBuild != 0 || stats.CFGBuild != 0 {
					t.Fatalf("discovery built procedure artifacts: %+v", stats)
				}
				published := diagnosticsByCode(analyzer.Diagnostics(doc), "VB001")
				if !reflect.DeepEqual(actions[0].Diagnostics, published) {
					t.Fatalf("associated = %+v, published = %+v", actions[0].Diagnostics, published)
				}
				got := applyQuickFixForTest(t, doc, actions[0])
				want := strings.ReplaceAll(tc.preamble+"Option Explicit\n"+tc.body, "\n", eol)
				if tc.name == "empty" {
					want = "Option Explicit\n" // No existing EOL to infer.
				}
				if got != want {
					t.Fatalf("corrected = %q, want %q", got, want)
				}
				doc = Document{Path: doc.Path, Source: got}
				if diagnostics := diagnosticsByCode(analyzer.Diagnostics(doc), "VB001"); len(diagnostics) != 0 {
					t.Fatalf("corrected diagnostics = %+v", diagnostics)
				}
				actions, err = analyzer.CodeActionsContext(t.Context(), doc, Range{}, []string{"quickfix"})
				if err != nil || len(actions) != 0 {
					t.Fatalf("second correction = %+v, err = %v", actions, err)
				}
			})
		}
	}
}

func TestOptionExplicitQuickFixWithholdsAmbiguousOrInvisibleEdits(t *testing.T) {
	for _, source := range []string{
		"Option Explicit\nSub Work()\nEnd Sub\n",
		"Option _\n Explicit\nSub Work()\nEnd Sub\n",
		"#If False Then\nOption Explicit\n#End If\nSub Work()\nEnd Sub\n",
		"#If VBA7 Then\nOption Explicit\n#End If\nSub Work()\nEnd Sub\n",
		"' xlflow:disable-line VB001\nSub Work()\nEnd Sub\n",
		"Sub Work(\nEnd Sub\n",
		"Option Compare Text\nAttribute VB_Name = \"Main\"\n",
	} {
		t.Run(source, func(t *testing.T) {
			analyzer := newTestAnalyzer(t)
			doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: source}
			actions, err := analyzer.CodeActionsContext(t.Context(), doc, Range{}, []string{"quickfix"})
			if err != nil || len(actions) != 0 {
				t.Fatalf("actions = %+v, err = %v", actions, err)
			}
		})
	}
}

func TestOptionExplicitQuickFixKeepsEncodingPolicy(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: "\ufeffSub Work()\nEnd Sub\n"}
	actions, err := analyzer.CodeActionsContext(t.Context(), doc, Range{}, []string{"quickfix"})
	if err == nil || len(actions) != 0 {
		t.Fatalf("BOM input actions = %+v, err = %v; want existing encoding rejection", actions, err)
	}
}

func TestLocalQuickFixMatchesVBEFixtures(t *testing.T) {
	root := filepath.Join("..", "..", "..", "testdata", "vbe-oracle", "cases")
	before, err := os.ReadFile(filepath.Join(root, "quickfix-local-before", "Main.bas"))
	if err != nil {
		t.Fatal(err)
	}
	after, err := os.ReadFile(filepath.Join(root, "quickfix-local-after", "Main.bas"))
	if err != nil {
		t.Fatal(err)
	}
	analyzer := newTestAnalyzer(t)
	analyzer.Config.Lint.DetectLetAssignment = true
	analyzer.Config.Lint.DetectRedundantOptionBaseZero = true
	doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: string(before)}
	actions, err := analyzer.CodeActionsContext(t.Context(), doc, Range{End: positionForByteOffset(doc.Source, len(doc.Source))}, []string{"quickfix"})
	if err != nil || len(actions) != 3 {
		t.Fatalf("actions = %+v, err = %v", actions, err)
	}
	var combined CodeAction
	for _, action := range actions {
		combined.Edits = append(combined.Edits, action.Edits...)
	}
	if got := applyQuickFixForTest(t, doc, combined); got != string(after) {
		t.Fatalf("corrected = %q, want VBE fixture %q", got, after)
	}
}

func TestLocalQuickFixPreservesSourceAndPublishedDiagnostic(t *testing.T) {
	for _, tc := range []struct{ name, code, source, want string }{
		{"let", "VB080", "Sub Work()\nDim x As Long\n    Let x = 1 ' 日本語😀\nEnd Sub\n", "Sub Work()\nDim x As Long\n    x = 1 ' 日本語😀\nEnd Sub\n"},
		{"let tabs", "VB080", "Sub Work()\nDim x As Long\n\tlEt\t x = 1\nEnd Sub\n", "Sub Work()\nDim x As Long\n\tx = 1\nEnd Sub\n"},
		{"let colon unicode", "VB080", "Sub Work()\nDim x As Long\nDebug.Print \"😀\": Let x = 1\nEnd Sub\n", "Sub Work()\nDim x As Long\nDebug.Print \"😀\": x = 1\nEnd Sub\n"},
		{"let continued value", "VB080", "Sub Work()\nDim x As Long\nLet x = _\n 1\nEnd Sub\n", "Sub Work()\nDim x As Long\nx = _\n 1\nEnd Sub\n"},
		{"base comment", "VB086", "  Option Base 0 ' 日本語😀\nSub Work()\nEnd Sub\n", "   ' 日本語😀\nSub Work()\nEnd Sub\n"},
		{"base tabs", "VB086", "oPtIoN\tBase\t0\nSub Work()\nEnd Sub\n", "\nSub Work()\nEnd Sub\n"},
	} {
		for _, eol := range []string{"\n", "\r\n"} {
			t.Run(tc.name+"/"+strings.ReplaceAll(eol, "\n", "LF"), func(t *testing.T) {
				analyzer := newTestAnalyzer(t)
				analyzer.Config.Lint.RequireOptionExplicit = false
				analyzer.Config.Lint.DetectLetAssignment = tc.code == "VB080"
				analyzer.Config.Lint.DetectRedundantOptionBaseZero = tc.code == "VB086"
				doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: strings.ReplaceAll(tc.source, "\n", eol)}
				snapshot := NewAnalysisSnapshot(doc)
				defer snapshot.Retire()
				doc = snapshot.Document()
				actions, err := analyzer.CodeActionsContext(t.Context(), doc, Range{End: positionForByteOffset(doc.Source, len(doc.Source))}, []string{"quickfix"})
				if err != nil || len(actions) != 1 {
					t.Fatalf("actions = %+v, err = %v", actions, err)
					return
				}
				if stats := snapshot.ProcedureArtifactStats(); stats.IRBuild != 0 || stats.CFGBuild != 0 {
					t.Fatalf("discovery built procedure artifacts: %+v", stats)
				}
				diagnostics := analyzer.Diagnostics(doc)
				index := slices.IndexFunc(diagnostics, func(d Diagnostic) bool { return d.Code == tc.code })
				if index < 0 || !reflect.DeepEqual(actions[0].Diagnostics, []Diagnostic{diagnostics[index]}) {
					t.Fatalf("associated diagnostics = %+v, published = %+v", actions[0].Diagnostics, diagnostics)
				}
				got := applyQuickFixForTest(t, doc, actions[0])
				if want := strings.ReplaceAll(tc.want, "\n", eol); got != want {
					t.Fatalf("corrected = %q, want %q", got, want)
				}
				doc = Document{Path: doc.Path, Source: got}
				for _, diagnostic := range analyzer.Diagnostics(doc) {
					if diagnostic.Code == tc.code || diagnostic.Code == "VB014" {
						t.Fatalf("unexpected corrected diagnostic: %+v", diagnostic)
					}
				}
				actions, err = analyzer.CodeActionsContext(t.Context(), doc, Range{End: positionForByteOffset(got, len(got))}, []string{"quickfix"})
				if err != nil || len(actions) != 0 {
					t.Fatalf("second correction = %+v, err = %v", actions, err)
				}
			})
		}
	}
}

func TestLocalQuickFixWithholdsUnsafeOrInvisibleEdits(t *testing.T) {
	for _, tc := range []struct {
		name, source string
		disabled     bool
	}{
		{"disabled", "Option Base 0\nSub Work()\nLet x = 1\nEnd Sub\n", true},
		{"continued Let", "Sub Work()\nLet _\n x = 1\nEnd Sub\n", false},
		{"continued base", "Option Base _\n 0\nSub Work()\nEnd Sub\n", false},
		{"base one", "Option Base 1\nSub Work()\nEnd Sub\n", false},
		{"implicit assignment", "Sub Work()\nx = 1\nEnd Sub\n", false},
		{"malformed", "Sub Work()\nLet = 1\nEnd Sub\n", false},
		{"suppressed let", "Sub Work()\nLet x = 1 ' xlflow:disable-line VB080\nEnd Sub\n", false},
		{"suppressed base", "' xlflow:disable-next-line VB086\nOption Base 0\nSub Work()\nEnd Sub\n", false},
	} {
		t.Run(tc.name, func(t *testing.T) {
			analyzer := newTestAnalyzer(t)
			analyzer.Config.Lint = config.LintConfig{DetectLetAssignment: !tc.disabled, DetectRedundantOptionBaseZero: !tc.disabled}
			doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: tc.source}
			snapshot := NewAnalysisSnapshot(doc)
			defer snapshot.Retire()
			doc = snapshot.Document()
			actions, err := analyzer.CodeActionsContext(t.Context(), doc, Range{End: positionForByteOffset(tc.source, len(tc.source))}, []string{"quickfix"})
			if err != nil || len(actions) != 0 {
				t.Fatalf("actions = %+v, err = %v", actions, err)
			}
			if tc.disabled && snapshot.ParseCount() != 0 {
				t.Fatal("disabled corrections parsed the document")
			}
		})
	}
}
