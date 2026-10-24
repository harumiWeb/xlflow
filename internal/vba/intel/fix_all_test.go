package intel

import (
	"path/filepath"
	"slices"
	"strings"
	"testing"
)

func TestDocumentFixAll(t *testing.T) {
	wantKinds := []string{
		"source.fixAll.xlflow.vb044",
		"source.fixAll.xlflow.vb080",
		"source.fixAll.xlflow.vb084",
		"source.fixAll.xlflow.vb086",
	}
	gotKinds := FixAllCodeActionKinds()
	if !slices.Equal(gotKinds, wantKinds) {
		t.Fatalf("FixAllCodeActionKinds() = %v, want %v", gotKinds, wantKinds)
	}
	gotKinds[0] = "mutated"
	if got := FixAllCodeActionKinds(); !slices.Equal(got, wantKinds) {
		t.Fatalf("FixAllCodeActionKinds() reused caller storage: %v", got)
	}

	wantCounts := map[string]int{"VB044": 2, "VB080": 2, "VB084": 2, "VB086": 2}
	for _, eol := range []string{"\n", "\r\n"} {
		name := "LF"
		if eol == "\r\n" {
			name = "CRLF"
		}
		t.Run(name, func(t *testing.T) {
			analyzer := newTestAnalyzer(t)
			enableAllDocumentFixAllRules(&analyzer)
			source := documentFixAllFixture(eol)
			doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: source}
			snapshot := NewAnalysisSnapshot(doc)
			defer snapshot.Retire()
			doc = snapshot.Document()

			actions, err := analyzer.CodeActionsContext(t.Context(), doc, Range{}, []string{"source.fixAll"})
			if err != nil {
				t.Fatal(err)
			}
			if len(actions) != len(wantKinds) {
				t.Fatalf("Fix All actions = %+v, want %d", actions, len(wantKinds))
			}
			for i, action := range actions {
				if action.Kind != wantKinds[i] {
					t.Fatalf("action[%d].Kind = %q, want %q", i, action.Kind, wantKinds[i])
				}
			}
			if stats := snapshot.ProcedureArtifactStats(); stats.IRBuild != 0 || stats.CFGBuild != 0 {
				t.Fatalf("Fix All built procedure artifacts: %+v", stats)
			}
			if builds, hits := snapshot.InteractiveIndexStats(); builds != 0 || hits != 0 {
				t.Fatalf("Fix All used the interactive index: builds=%d hits=%d", builds, hits)
			}

			fullSelection := Range{End: positionForDocumentByteOffset(doc, len(doc.Source))}
			quickFixActions, err := analyzer.CodeActionsContext(t.Context(), doc, fullSelection, []string{"quickfix"})
			if err != nil {
				t.Fatal(err)
			}

			publishedDoc := doc
			publishedDoc.Snapshot = nil
			published := analyzer.Diagnostics(publishedDoc)
			for _, action := range actions {
				code := strings.ToUpper(strings.TrimPrefix(action.Kind, "source.fixAll.xlflow."))
				wantDiagnostics := diagnosticsByCode(published, code)
				if len(wantDiagnostics) != wantCounts[code] {
					t.Fatalf("published %s diagnostics = %+v, want %d", code, wantDiagnostics, wantCounts[code])
				}
				if !slices.Equal(action.Diagnostics, wantDiagnostics) {
					t.Errorf("%s associated diagnostics = %+v, published = %+v", code, action.Diagnostics, wantDiagnostics)
				}
				if len(action.Edits) != wantCounts[code] {
					t.Errorf("%s edits = %+v, want %d", code, action.Edits, wantCounts[code])
				}
				var individualEdits []CodeActionEdit
				for _, quickFix := range quickFixActions {
					if len(quickFix.Diagnostics) == 1 && quickFix.Diagnostics[0].Code == code {
						individualEdits = append(individualEdits, quickFix.Edits...)
					}
				}
				if len(individualEdits) != len(action.Edits) {
					t.Errorf("%s individual quick-fix edits = %d, Fix All edits = %d", code, len(individualEdits), len(action.Edits))
				} else if individual := applyQuickFixForTest(t, doc, CodeAction{Edits: individualEdits}); individual != applyQuickFixForTest(t, doc, action) {
					t.Errorf("%s individual quick fixes produce %q, Fix All produces %q", code, individual, applyQuickFixForTest(t, doc, action))
				}

				corrected := applyQuickFixForTest(t, doc, action)
				if want := documentFixAllExpectedSource(code, source); corrected != want {
					t.Errorf("%s corrected source = %q, want exact bytes %q", code, corrected, want)
					continue
				}
				correctedDoc := Document{Path: doc.Path, Source: corrected}
				if remaining := diagnosticsByCode(analyzer.Diagnostics(correctedDoc), code); len(remaining) != 0 {
					t.Errorf("%s diagnostics remain after Fix All: %+v", code, remaining)
				}
				again, err := analyzer.CodeActionsContext(t.Context(), correctedDoc, Range{}, []string{action.Kind})
				if err != nil || len(again) != 0 {
					t.Errorf("second %s Fix All = %+v, err = %v; want none", code, again, err)
				}
			}
		})
	}
}

func TestDocumentFixAllKindFiltering(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	enableAllDocumentFixAllRules(&analyzer)
	source := documentFixAllFixture("\n")
	doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: source}
	selection := Range{End: positionForByteOffset(doc.Source, len(doc.Source))}
	for _, tc := range []struct {
		name         string
		only         []string
		wantFixAll   []string
		wantQuickFix bool
	}{
		{name: "source root parent", only: []string{"source"}, wantFixAll: FixAllCodeActionKinds()},
		{name: "Fix All parent", only: []string{"source.fixAll"}, wantFixAll: FixAllCodeActionKinds()},
		{name: "provider parent", only: []string{"source.fixAll.xlflow"}, wantFixAll: FixAllCodeActionKinds()},
		{name: "exact kind", only: []string{"source.fixAll.xlflow.vb084"}, wantFixAll: []string{"source.fixAll.xlflow.vb084"}},
		{name: "empty kind", only: []string{""}, wantQuickFix: true},
		{name: "quickfix kind", only: []string{"quickfix"}, wantQuickFix: true},
		{name: "empty kind and exact provider", only: []string{"", "source.fixAll.xlflow.vb084"}, wantFixAll: []string{"source.fixAll.xlflow.vb084"}, wantQuickFix: true},
		{name: "no only kinds", wantQuickFix: true},
	} {
		t.Run(tc.name, func(t *testing.T) {
			actions, err := analyzer.CodeActionsContext(t.Context(), doc, selection, tc.only)
			if err != nil {
				t.Fatal(err)
			}
			var gotFixAll []string
			hasQuickFix := false
			for _, action := range actions {
				if strings.HasPrefix(action.Kind, "source.fixAll") {
					gotFixAll = append(gotFixAll, action.Kind)
				}
				hasQuickFix = hasQuickFix || action.Kind == "quickfix"
			}
			if !slices.Equal(gotFixAll, tc.wantFixAll) {
				t.Fatalf("Fix All action kinds = %v, want %v (all actions: %+v)", gotFixAll, tc.wantFixAll, actions)
			}
			if hasQuickFix != tc.wantQuickFix {
				t.Fatalf("ordinary quick fixes present = %t, want %t (all actions: %+v)", hasQuickFix, tc.wantQuickFix, actions)
			}
		})
	}
}

func TestDocumentFixAllRespectsSuppressions(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	enableAllDocumentFixAllRules(&analyzer)
	source := `Option Base 0 ' xlflow:disable-line VB086
Public Sub Alpha()
    Const PROCEDURE_NAME As String = "Old" ' xlflow:disable-line VB044
    Dim count As Long
    Dim i As Long
    Let count = 1 ' xlflow:disable-line VB080
    For i = 1 To 3 Step 1 ' xlflow:disable-line VB084
        Debug.Print i
    Next i
End Sub
`
	doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: source}
	for _, code := range []string{"VB044", "VB080", "VB084", "VB086"} {
		if diagnostics := diagnosticsByCode(analyzer.Diagnostics(doc), code); len(diagnostics) != 0 {
			t.Fatalf("suppressed %s diagnostics = %+v, want none", code, diagnostics)
		}
	}
	actions, err := analyzer.CodeActionsContext(t.Context(), doc, Range{}, []string{"source.fixAll.xlflow"})
	if err != nil || len(actions) != 0 {
		t.Fatalf("Fix All actions for suppressed diagnostics = %+v, err = %v; want none", actions, err)
	}
}

func TestDocumentFixAllWithholdsUnsafeEdits(t *testing.T) {
	for _, tc := range []struct {
		name, code, source, want string
	}{
		{
			name: "continued Let keyword", code: "VB080",
			source: "Sub Work()\nDim safe As Long\nDim unsafe As Long\nLet safe = _\n 1\nLet _\n unsafe = 2\nEnd Sub\n",
			want:   "Sub Work()\nDim safe As Long\nDim unsafe As Long\nsafe = _\n 1\nLet _\n unsafe = 2\nEnd Sub\n",
		},
		{
			name: "continued Step keyword", code: "VB084",
			source: "Sub Work()\nDim i As Long\nDim j As Long\nFor i = 1 To 3 Step 1\n    Debug.Print i\nNext i\nFor j = 1 To 3 _\n    Step 1\n    Debug.Print j\nNext j\nEnd Sub\n",
			want:   "Sub Work()\nDim i As Long\nDim j As Long\nFor i = 1 To 3\n    Debug.Print i\nNext i\nFor j = 1 To 3 _\n    Step 1\n    Debug.Print j\nNext j\nEnd Sub\n",
		},
		{
			name: "continued Option Base", code: "VB086",
			source: "Option Base 0\nOption Base _\n 0\nSub Work()\nEnd Sub\n",
			want:   "\nOption Base _\n 0\nSub Work()\nEnd Sub\n",
		},
	} {
		t.Run(tc.name, func(t *testing.T) {
			analyzer := newTestAnalyzer(t)
			enableDocumentFixAllRule(&analyzer, tc.code)
			doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: tc.source}
			actions, err := analyzer.CodeActionsContext(t.Context(), doc, Range{}, []string{documentFixAllKind(tc.code)})
			if err != nil || len(actions) != 1 {
				t.Fatalf("Fix All actions = %+v, err = %v; want one safe aggregate", actions, err)
			}
			if len(actions[0].Diagnostics) != 1 || actions[0].Diagnostics[0].Code != tc.code {
				t.Fatalf("associated diagnostics = %+v, want one safe %s diagnostic", actions[0].Diagnostics, tc.code)
			}
			if got := applyQuickFixForTest(t, doc, actions[0]); got != tc.want {
				t.Fatalf("corrected source = %q, want safe edit only %q", got, tc.want)
			}
		})
	}
}

func TestDocumentFixAllSkipsParsingForDisabledAndUnknownRequests(t *testing.T) {
	for _, tc := range []struct {
		name  string
		only  []string
		setup func(*Analyzer)
	}{
		{
			name: "unknown kind",
			only: []string{"source.fixAll.xlflow.unknown"},
			setup: func(analyzer *Analyzer) {
				enableAllDocumentFixAllRules(analyzer)
			},
		},
		{
			name: "all providers disabled",
			only: []string{"source.fixAll.xlflow.vb084"},
			setup: func(analyzer *Analyzer) {
				analyzer.Config.Lint.RequireOptionExplicit = false
				analyzer.Config.Lint.ProcedureNameConstant.Enabled = false
				analyzer.Config.Lint.DetectLetAssignment = false
				analyzer.Config.Lint.DetectRedundantStepOne = false
				analyzer.Config.Lint.DetectRedundantOptionBaseZero = false
			},
		},
	} {
		t.Run(tc.name, func(t *testing.T) {
			analyzer := newTestAnalyzer(t)
			analyzer.SourceOnly = true
			analyzer.DB = nil
			tc.setup(&analyzer)
			doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: "Sub Work()\nEnd Sub\n"}
			snapshot := NewAnalysisSnapshot(doc)
			defer snapshot.Retire()
			doc = snapshot.Document()

			actions, err := analyzer.CodeActionsContext(t.Context(), doc, Range{}, tc.only)
			if err != nil || len(actions) != 0 {
				t.Fatalf("actions = %+v, err = %v; want none", actions, err)
			}
			if got := snapshot.ParseCount(); got != 0 {
				t.Fatalf("parsed document %d times for an unserved request, want zero", got)
			}
			if stats := snapshot.ProcedureArtifactStats(); stats.IRBuild != 0 || stats.CFGBuild != 0 {
				t.Fatalf("request built procedure artifacts: %+v", stats)
			}
			if builds, hits := snapshot.InteractiveIndexStats(); builds != 0 || hits != 0 {
				t.Fatalf("request used the interactive index: builds=%d hits=%d", builds, hits)
			}
		})
	}
}

func TestDocumentFixAllExcludesIneligibleRules(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	enableDocumentFixAllRule(&analyzer, "VB084")
	analyzer.Config.Lint.RequireOptionExplicit = true
	source := `Public Sub Work()
    Dim target As Excel.Worksheet
    Dim source As Excel.Worksheet
    Dim count As Long
    Dim i As Long
    target = source
    Set count = 1
    For i = 1 To 3 Step 1
        Debug.Print i
    Next i
End Sub
`
	doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: source}
	published := analyzer.Diagnostics(doc)
	for _, code := range []string{"VB001", "VB036", "VB037"} {
		if diagnostics := diagnosticsByCode(published, code); len(diagnostics) == 0 {
			t.Fatalf("control source has no %s diagnostic: %+v", code, published)
		}
	}

	actions, err := analyzer.CodeActionsContext(t.Context(), doc, Range{}, []string{"source.fixAll"})
	if err != nil || len(actions) != 1 || actions[0].Kind != "source.fixAll.xlflow.vb084" {
		t.Fatalf("Fix All actions = %+v, err = %v; want only VB084", actions, err)
	}
	for _, code := range []string{"VB001", "VB036", "VB037"} {
		actions, err := analyzer.CodeActionsContext(t.Context(), doc, Range{}, []string{documentFixAllKind(code)})
		if err != nil || len(actions) != 0 {
			t.Errorf("ineligible %s request actions = %+v, err = %v; want none", code, actions, err)
		}
	}
}

func TestAggregateDocumentFixes(t *testing.T) {
	doc := Document{Source: "abcdef"}
	fix := func(code string, start, end int) CodeAction {
		return CodeAction{
			Diagnostics: []Diagnostic{{Code: code}},
			Edits: []CodeActionEdit{{Range: Range{
				Start: positionForByteOffset(doc.Source, start),
				End:   positionForByteOffset(doc.Source, end),
			}}},
		}
	}

	action, ok := aggregateDocumentFixes(doc, "VB084", []CodeAction{
		fix("VB084", 4, 6),
		fix("VB080", 0, 2),
		fix("VB084", 0, 2),
		fix("VB084", 2, 4),
	})
	if !ok {
		t.Fatal("aggregateDocumentFixes rejected sorted adjacent edits")
	}
	if action.Kind != "source.fixAll.xlflow.vb084" || action.Title != "Fix all VB084 occurrences in this document" {
		t.Fatalf("aggregate action metadata = %+v", action)
	}
	if len(action.Diagnostics) != 3 {
		t.Fatalf("aggregated diagnostics = %+v, want only three target-rule diagnostics", action.Diagnostics)
	}
	if len(action.Edits) != 3 {
		t.Fatalf("aggregated edits = %+v, want three adjacent edits", action.Edits)
	}
	for i, wantStart := range []int{0, 2, 4} {
		if got := byteOffsetForDocumentPosition(doc, action.Edits[i].Range.Start); got != wantStart {
			t.Errorf("edit[%d] starts at byte %d, want %d", i, got, wantStart)
		}
	}
	if _, ok := aggregateDocumentFixes(doc, "VB084", []CodeAction{fix("VB080", 0, 1)}); ok {
		t.Fatal("aggregation succeeded without a target-rule fix")
	}
	if _, ok := aggregateDocumentFixes(doc, "VB001", []CodeAction{fix("VB001", 0, 1)}); ok {
		t.Fatal("aggregation accepted an ineligible rule")
	}

	for _, tc := range []struct {
		name  string
		doc   Document
		fixes []CodeAction
	}{
		{name: "reversed range", doc: doc, fixes: []CodeAction{{Diagnostics: []Diagnostic{{Code: "VB084"}}, Edits: []CodeActionEdit{{Range: Range{Start: Position{Character: 2}, End: Position{Character: 1}}}}}}},
		{name: "negative position", doc: doc, fixes: []CodeAction{{Diagnostics: []Diagnostic{{Code: "VB084"}}, Edits: []CodeActionEdit{{Range: Range{Start: Position{Line: -1}, End: Position{}}}}}}},
		{name: "position past document", doc: doc, fixes: []CodeAction{{Diagnostics: []Diagnostic{{Code: "VB084"}}, Edits: []CodeActionEdit{{Range: Range{Start: Position{Character: 0}, End: Position{Character: 99}}}}}}},
		{name: "overlapping ranges", doc: doc, fixes: []CodeAction{fix("VB084", 0, 3), fix("VB084", 2, 4)}},
		{name: "duplicate edits", doc: doc, fixes: []CodeAction{fix("VB084", 1, 3), fix("VB084", 1, 3)}},
		{name: "same-point insertions", doc: doc, fixes: []CodeAction{fix("VB084", 2, 2), fix("VB084", 2, 2)}},
		{name: "invalid surrogate start", doc: Document{Source: "a😀b"}, fixes: []CodeAction{{Diagnostics: []Diagnostic{{Code: "VB084"}}, Edits: []CodeActionEdit{{Range: Range{Start: Position{Character: 2}, End: Position{Character: 3}}}}}}},
		{name: "invalid surrogate end", doc: Document{Source: "a😀b"}, fixes: []CodeAction{{Diagnostics: []Diagnostic{{Code: "VB084"}}, Edits: []CodeActionEdit{{Range: Range{Start: Position{Character: 1}, End: Position{Character: 2}}}}}}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			if action, ok := aggregateDocumentFixes(tc.doc, "VB084", tc.fixes); ok {
				t.Fatalf("aggregateDocumentFixes accepted invalid edits: %+v", action)
			}
		})
	}
}

func enableAllDocumentFixAllRules(analyzer *Analyzer) {
	analyzer.Config.Lint.RequireOptionExplicit = false
	analyzer.Config.Lint.ProcedureNameConstant.Enabled = true
	analyzer.Config.Lint.ProcedureNameConstant.ConstantName = "PROCEDURE_NAME"
	analyzer.Config.Lint.DetectLetAssignment = true
	analyzer.Config.Lint.DetectRedundantStepOne = true
	analyzer.Config.Lint.DetectRedundantOptionBaseZero = true
}

func enableDocumentFixAllRule(analyzer *Analyzer, code string) {
	analyzer.Config.Lint.RequireOptionExplicit = false
	switch code {
	case "VB044":
		analyzer.Config.Lint.ProcedureNameConstant.Enabled = true
		analyzer.Config.Lint.ProcedureNameConstant.ConstantName = "PROCEDURE_NAME"
	case "VB080":
		analyzer.Config.Lint.DetectLetAssignment = true
	case "VB084":
		analyzer.Config.Lint.DetectRedundantStepOne = true
	case "VB086":
		analyzer.Config.Lint.DetectRedundantOptionBaseZero = true
	default:
		panic("unsupported Fix All test rule: " + code)
	}
}

func documentFixAllFixture(eol string) string {
	source := `' module trivia: 日本語😀
    Option Base 0 ' keep first option 😀
	Option Base 0 ' keep second option 日本語
Public Sub Alpha()
	Const PROCEDURE_NAME As String = "旧" ' keep first literal 😀
    Dim first As Long, second As Long, i As Long, j As Long
	Let first = 1 ' keep first assignment 日本語
    Let second = 2 ' keep second assignment 😀
    For i = 1 To 3 Step 1 ' keep first loop 日本語
        Debug.Print i
    Next i
    For j = 1 To 3 Step 1 ' keep second loop 😀
        Debug.Print j
    Next j
End Sub
Public Function Beta() As Long
    Const PROCEDURE_NAME As String = "古い名前" ' keep second literal 日本語
    Beta = 1
End Function
`
	return strings.ReplaceAll(source, "\n", eol)
}

func documentFixAllExpectedSource(code, source string) string {
	switch code {
	case "VB044":
		source = strings.Replace(source, `"旧"`, `"Alpha"`, 1)
		return strings.Replace(source, `"古い名前"`, `"Beta"`, 1)
	case "VB080":
		source = strings.Replace(source, "\tLet first", "\tfirst", 1)
		return strings.Replace(source, "    Let second", "    second", 1)
	case "VB084":
		return strings.ReplaceAll(source, " Step 1", "")
	case "VB086":
		return strings.ReplaceAll(source, "Option Base 0", "")
	default:
		panic("unsupported Fix All expected-source rule: " + code)
	}
}
