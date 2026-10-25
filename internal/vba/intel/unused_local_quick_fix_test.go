package intel

import (
	"os"
	"path/filepath"
	"reflect"
	"slices"
	"strconv"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestUnusedLocalQuickFixPreservesSourceAndPublishedDiagnostic(t *testing.T) {
	types := []string{
		"Boolean", "Byte", "Integer", "Long", "LongLong", "LongPtr",
		"Single", "Double", "Currency", "Date", "String", "Variant",
	}
	for _, typeName := range types {
		for _, eol := range []string{"\n", "\r\n"} {
			eolName := "LF"
			if eol == "\r\n" {
				eolName = "CRLF"
			}
			t.Run(typeName+"/"+eolName, func(t *testing.T) {
				analyzer := newTestAnalyzer(t)
				analyzer.DB = nil
				analyzer.SourceOnly = true
				source := strings.ReplaceAll(
					"Option Explicit\nSub Work()\n    Dim scratch As "+typeName+" ' keep 日本語😀\nEnd Sub\n",
					"\n",
					eol,
				)
				doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: source}
				snapshot := NewAnalysisSnapshot(doc)
				defer snapshot.Retire()
				doc = snapshot.Document()

				actions := requestUnusedLocalQuickFixes(t, analyzer, doc)
				if len(actions) != 1 {
					t.Fatalf("quick fixes = %+v, want one VB020 action", actions)
				}
				action := actions[0]
				if action.Title != "Remove unused local variable" || action.Kind != "quickfix" {
					t.Fatalf("action = %+v, want the individual VB020 quick fix", action)
				}
				if len(action.Diagnostics) != 1 || action.Diagnostics[0].Code != "VB020" {
					t.Fatalf("associated diagnostics = %+v, want one VB020", action.Diagnostics)
				}
				if len(action.Edits) != 1 || action.Edits[0].NewText != "" {
					t.Fatalf("edits = %+v, want one deletion", action.Edits)
				}
				wantRange := Range{
					Start: Position{Line: 2, Character: 4},
					End:   Position{Line: 2, Character: 4 + len("Dim scratch As "+typeName)},
				}
				if action.Edits[0].Range != wantRange {
					t.Fatalf("edit range = %+v, want statement span %+v", action.Edits[0].Range, wantRange)
				}

				assertUnusedLocalQuickFixSnapshotStayedLight(t, snapshot)

				published := diagnosticsByCode(analyzer.Diagnostics(doc), "VB020")
				if len(published) != 1 || !reflect.DeepEqual(action.Diagnostics, published) {
					t.Fatalf("associated diagnostics = %+v, published VB020 = %+v", action.Diagnostics, published)
				}

				got, err := ApplyCodeActionEdits(t.Context(), doc, action.Edits)
				if err != nil {
					t.Fatal(err)
				}
				want := strings.ReplaceAll(
					"Option Explicit\nSub Work()\n     ' keep 日本語😀\nEnd Sub\n",
					"\n",
					eol,
				)
				if got != want {
					t.Fatalf("corrected source = %q, want exact preserved bytes %q", got, want)
				}
			})
		}
	}
}

func TestUnusedLocalQuickFixWithholdsSuppressedDisabledAndUnsafeCases(t *testing.T) {
	for _, tc := range []struct {
		name        string
		source      string
		configure   func(*Analyzer)
		wantVB020   bool
		wantNoParse bool
	}{
		{
			name:   "suppressed",
			source: "Option Explicit\nSub Work()\n    Dim scratch As Long ' xlflow:disable-line VB020\nEnd Sub\n",
		},
		{
			name:   "disabled",
			source: "Option Explicit\nSub Work()\n    Dim scratch As Long\nEnd Sub\n",
			configure: func(analyzer *Analyzer) {
				analyzer.Config.Lint = config.LintConfig{}
			},
			wantNoParse: true,
		},
		{
			name:      "write only",
			source:    "Option Explicit\nSub Work()\n    Dim scratch As Long\n    scratch = 1\nEnd Sub\n",
			wantVB020: true,
		},
		{
			name:      "static local",
			source:    "Option Explicit\nSub Work()\n    Static scratch As Long\nEnd Sub\n",
			wantVB020: true,
		},
		{
			name:      "static procedure",
			source:    "Option Explicit\nStatic Sub Work()\n    Dim scratch As Long\nEnd Sub\n",
			wantVB020: true,
		},
		{
			name:   "ignored unused prefix",
			source: "Option Explicit\nSub Work()\n    Dim unusedScratch As Long\nEnd Sub\n",
		},
		{
			name:   "ignored ignore prefix",
			source: "Option Explicit\nSub Work()\n    Dim ignoreScratch As Long\nEnd Sub\n",
		},
	} {
		t.Run(tc.name, func(t *testing.T) {
			analyzer := newTestAnalyzer(t)
			analyzer.DB = nil
			analyzer.SourceOnly = true
			if tc.configure != nil {
				tc.configure(&analyzer)
			}
			doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: tc.source}
			snapshot := NewAnalysisSnapshot(doc)
			defer snapshot.Retire()
			doc = snapshot.Document()

			actions := requestUnusedLocalQuickFixes(t, analyzer, doc)
			if len(actions) != 0 {
				t.Fatalf("quick fixes = %+v, want none", actions)
			}
			if tc.wantNoParse && snapshot.ParseCount() != 0 {
				t.Fatalf("disabled VB020 parsed the snapshot %d times, want zero", snapshot.ParseCount())
			}
			if tc.wantVB020 {
				if diagnostics := diagnosticsByCode(analyzer.Diagnostics(doc), "VB020"); len(diagnostics) != 1 {
					t.Fatalf("VB020 diagnostics = %+v, want one diagnostic without a deletion action", diagnostics)
				}
			} else if tc.name == "suppressed" {
				if diagnostics := diagnosticsByCode(analyzer.Diagnostics(doc), "VB020"); len(diagnostics) != 0 {
					t.Fatalf("suppressed VB020 diagnostics = %+v, want none", diagnostics)
				}
			}
		})
	}
}

func TestUnusedLocalQuickFixDoesNotExpandDocumentFixAll(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	analyzer.Config.Lint.DetectUnusedLocalVariables = true
	enableAllDocumentFixAllRules(&analyzer)
	wantKinds := []string{
		"source.fixAll.xlflow.vb044",
		"source.fixAll.xlflow.vb080",
		"source.fixAll.xlflow.vb084",
		"source.fixAll.xlflow.vb086",
	}
	if got := FixAllCodeActionKinds(); !slices.Equal(got, wantKinds) {
		t.Fatalf("FixAllCodeActionKinds() = %v, want existing allowlist %v", got, wantKinds)
	}

	source := documentFixAllFixture("\n")
	source = strings.Replace(source, "Public Sub Alpha()\n", "Public Sub Alpha()\n    Dim scratch As Long\n", 1)
	doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: source}
	selection := Range{End: positionForByteOffset(doc.Source, len(doc.Source))}
	actions, err := analyzer.CodeActionsContext(t.Context(), doc, selection, []string{"source.fixAll.xlflow"})
	if err != nil {
		t.Fatal(err)
	}
	var gotKinds []string
	for _, action := range actions {
		if strings.HasPrefix(action.Kind, "source.fixAll.") {
			gotKinds = append(gotKinds, action.Kind)
		}
		for _, diagnostic := range action.Diagnostics {
			if diagnostic.Code == "VB020" {
				t.Fatalf("parent Fix All returned a VB020 action: %+v", action)
			}
		}
	}
	if !slices.Equal(gotKinds, wantKinds) {
		t.Fatalf("parent Fix All kinds = %v, want existing four rules %v", gotKinds, wantKinds)
	}

	actions, err = analyzer.CodeActionsContext(t.Context(), doc, selection, []string{"source.fixAll.xlflow.vb020"})
	if err != nil {
		t.Fatal(err)
	}
	if len(actions) != 0 {
		t.Fatalf("VB020 exact Fix All request returned actions: %+v", actions)
	}
}

func TestUnusedLocalQuickFixScalesAcrossUnrelatedProcedures(t *testing.T) {
	var source strings.Builder
	source.WriteString("Option Explicit\n")
	for i := range 128 {
		name := strconv.Itoa(i)
		source.WriteString("Sub Unrelated" + name + "()\n")
		source.WriteString("    Dim localValue As Long\n")
		source.WriteString("    localValue = 1\n")
		source.WriteString("    localValue = localValue + 1\n")
		source.WriteString("    Debug.Print localValue\n")
		source.WriteString("End Sub\n")
	}
	source.WriteString("Sub Eligible()\n")
	eligibleLine := strings.Count(source.String(), "\n")
	source.WriteString("    Dim scratch As Long ' eligible\n")
	source.WriteString("End Sub\n")

	analyzer := newTestAnalyzer(t)
	analyzer.DB = nil
	analyzer.SourceOnly = true
	doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: source.String()}
	snapshot := NewAnalysisSnapshot(doc)
	defer snapshot.Retire()
	doc = snapshot.Document()

	actions := requestUnusedLocalQuickFixes(t, analyzer, doc)
	if len(actions) != 1 {
		t.Fatalf("quick fixes = %+v, want one action for the eligible declaration", actions)
	}
	action := actions[0]
	if action.Title != "Remove unused local variable" || len(action.Diagnostics) != 1 || action.Diagnostics[0].Code != "VB020" {
		t.Fatalf("action = %+v, want the individual VB020 fix", action)
	}
	if len(action.Edits) != 1 || action.Edits[0].Range.Start.Line != eligibleLine || action.Edits[0].Range.Start.Character != 4 {
		t.Fatalf("edits = %+v, want the standalone Dim in procedure line %d", action.Edits, eligibleLine)
	}
	assertUnusedLocalQuickFixSnapshotStayedLight(t, snapshot)

	published := diagnosticsByCode(analyzer.Diagnostics(doc), "VB020")
	if len(published) != 1 || !reflect.DeepEqual(action.Diagnostics, published) {
		t.Fatalf("associated diagnostics = %+v, published VB020 = %+v", action.Diagnostics, published)
	}
}

func TestUnusedLocalQuickFixMatchesVBEFixtures(t *testing.T) {
	cases := filepath.Join("..", "..", "..", "testdata", "vbe-oracle", "cases")
	before, err := os.ReadFile(filepath.Join(cases, "quickfix-unused-local-before", "Main.bas"))
	if err != nil {
		t.Fatal(err)
	}
	after, err := os.ReadFile(filepath.Join(cases, "quickfix-unused-local-after", "Main.bas"))
	if err != nil {
		t.Fatal(err)
	}

	analyzer := newTestAnalyzer(t)
	analyzer.DB = nil
	analyzer.SourceOnly = true
	doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: string(before)}
	snapshot := NewAnalysisSnapshot(doc)
	defer snapshot.Retire()
	doc = snapshot.Document()
	actions := requestUnusedLocalQuickFixes(t, analyzer, doc)
	var unusedActions []CodeAction
	for _, action := range actions {
		if action.Title == "Remove unused local variable" && len(action.Diagnostics) == 1 && action.Diagnostics[0].Code == "VB020" {
			unusedActions = append(unusedActions, action)
		}
	}
	if len(unusedActions) != 1 {
		t.Fatalf("VBE before fixture actions = %+v, want one VB020 removal", actions)
	}
	assertUnusedLocalQuickFixSnapshotStayedLight(t, snapshot)

	published := diagnosticsByCode(analyzer.Diagnostics(doc), "VB020")
	if len(published) != 1 || !reflect.DeepEqual(unusedActions[0].Diagnostics, published) {
		t.Fatalf("fixture action diagnostics = %+v, published VB020 = %+v", unusedActions[0].Diagnostics, published)
	}
	got, err := ApplyCodeActionEdits(t.Context(), doc, unusedActions[0].Edits)
	if err != nil {
		t.Fatal(err)
	}
	if got != string(after) {
		t.Fatalf("applying VB020 fix produced %q, want promoted VBE after fixture %q", got, after)
	}
}

func requestUnusedLocalQuickFixes(t *testing.T, analyzer Analyzer, doc Document) []CodeAction {
	t.Helper()
	actions, err := analyzer.CodeActionsContext(
		t.Context(),
		doc,
		Range{End: positionForByteOffset(doc.Source, len(doc.Source))},
		[]string{"quickfix"},
	)
	if err != nil {
		t.Fatal(err)
	}
	return actions
}

func assertUnusedLocalQuickFixSnapshotStayedLight(t *testing.T, snapshot *AnalysisSnapshot) {
	t.Helper()
	if stats := snapshot.ProcedureArtifactStats(); stats.IRBuild != 0 || stats.CFGBuild != 0 {
		t.Fatalf("quick-fix discovery built procedure artifacts: %+v", stats)
	}
	if builds, hits := snapshot.InteractiveIndexStats(); builds != 0 || hits != 0 {
		t.Fatalf("quick-fix discovery used the interactive index: builds=%d hits=%d", builds, hits)
	}
	snapshot.symbolsMu.Lock()
	symbolsBuilt := snapshot.symbolsDone
	snapshot.symbolsMu.Unlock()
	if symbolsBuilt || snapshot.index != nil || snapshot.indexErr != nil || snapshot.semanticIdentifiers != nil {
		t.Fatalf(
			"quick-fix discovery populated a snapshot index: symbols=%v document=%v semantic=%v",
			symbolsBuilt,
			snapshot.index != nil || snapshot.indexErr != nil,
			snapshot.semanticIdentifiers != nil,
		)
	}
}
