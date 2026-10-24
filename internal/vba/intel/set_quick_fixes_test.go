package intel

import (
	"os"
	"path/filepath"
	"reflect"
	"strconv"
	"strings"
	"testing"
)

func TestSetAssignmentQuickFixPreservesSourceAndPublishedDiagnostic(t *testing.T) {
	tests := []struct {
		name, source, want, eol string
		code, title, editAnchor string
		editOffset, editLength  int
		newText                 string
		customTypes             bool
	}{
		{
			name:   "local object references with tabs and Unicode comment",
			source: "' 前置コメント 😀\nOption Explicit\nPublic Sub Work()\n\tDim target As Excel.Worksheet\n\tDim source As Excel.Worksheet\n\ttarget\t=\tsource ' 日本語 😀\nEnd Sub\n",
			want:   "' 前置コメント 😀\nOption Explicit\nPublic Sub Work()\n\tDim target As Excel.Worksheet\n\tDim source As Excel.Worksheet\n\tSet target\t=\tsource ' 日本語 😀\nEnd Sub\n",
			eol:    "\n", code: "VB036", title: "Insert Set for object assignment",
			editAnchor: "target\t=\tsource", newText: "Set ",
		},
		{
			name:   "parameter references replace Let with custom assignability",
			source: "Option Explicit\nPublic Sub Work(ByVal target As Test.Control, ByVal source As Test.TextBox)\n\tLet target = source ' 引数 😀\nEnd Sub\n",
			want:   "Option Explicit\nPublic Sub Work(ByVal target As Test.Control, ByVal source As Test.TextBox)\n\tSet target = source ' 引数 😀\nEnd Sub\n",
			eol:    "\r\n", code: "VB036", title: "Insert Set for object assignment",
			editAnchor: "Let target", editLength: len("Let "), newText: "Set ", customTypes: true,
		},
		{
			name:   "module object references",
			source: "Option Explicit\nPrivate target As Excel.Worksheet\nPrivate source As Excel.Worksheet\nPublic Sub Work()\n\ttarget = source\nEnd Sub\n",
			want:   "Option Explicit\nPrivate target As Excel.Worksheet\nPrivate source As Excel.Worksheet\nPublic Sub Work()\n\tSet target = source\nEnd Sub\n",
			eol:    "\n", code: "VB036", title: "Insert Set for object assignment",
			editAnchor: "target = source", newText: "Set ",
		},
		{
			name:   "scalar literal removal with tabs and Unicode comment",
			source: "Option Explicit\nPublic Sub Work()\n\tDim count As Long\n\tSet count\t=\t1 ' 値 😀\nEnd Sub\n",
			want:   "Option Explicit\nPublic Sub Work()\n\tDim count As Long\n\tcount\t=\t1 ' 値 😀\nEnd Sub\n",
			eol:    "\r\n", code: "VB037", title: "Remove Set from value assignment",
			editAnchor: "Set count", editLength: len("Set "),
		},
		{
			name:   "scalar typed identifier removal",
			source: "Option Explicit\nPublic Sub Work()\n\tDim count As Long\n\tDim other As Long\n\tSet count = other ' scalar 😀\nEnd Sub\n",
			want:   "Option Explicit\nPublic Sub Work()\n\tDim count As Long\n\tDim other As Long\n\tcount = other ' scalar 😀\nEnd Sub\n",
			eol:    "\n", code: "VB037", title: "Remove Set from value assignment",
			editAnchor: "Set count", editLength: len("Set "),
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			analyzer := newTestAnalyzer(t)
			if tc.customTypes {
				addSetAssignmentTestTypes(t, &analyzer)
				if assignable, known := analyzer.DB.IsAssignable("Test.Control", "Test.TextBox"); !known || !assignable {
					t.Fatalf("test type relation = %v, %v; want known assignability", assignable, known)
				}
			}
			source := strings.ReplaceAll(tc.source, "\n", tc.eol)
			want := strings.ReplaceAll(tc.want, "\n", tc.eol)
			assertSetAssignmentQuickFix(t, analyzer, source, want, tc.code, tc.title, tc.editAnchor, tc.editOffset, tc.editLength, tc.newText)
		})
	}
}

func TestSetAssignmentQuickFixWithholdsUnsafeCorrections(t *testing.T) {
	tests := []struct {
		name, source string
		customTypes  bool
		relation     string
	}{
		{
			name:   "compound string expression is not a scalar literal",
			source: "Option Explicit\nSub Work()\nDim target As String\nDim source As Excel.Range\nSet target = \"a\" & source & \"b\"\nEnd Sub\n",
		},
		{
			name:   "colon assignment",
			source: "Option Explicit\nSub Work()\nDim target As Excel.Worksheet\nDim source As Excel.Worksheet\ntarget = source: Debug.Print 1\nEnd Sub\n",
		},
		{
			name:   "line label",
			source: "Option Explicit\nSub Work()\nDim target As Excel.Worksheet\nDim source As Excel.Worksheet\nAgain: target = source\nEnd Sub\n",
		},
		{
			name:   "malformed whole document",
			source: "Option Explicit\nPublic Sub Work()\nDim target As Excel.Worksheet\nDim source As Excel.Worksheet\ntarget = source\n",
		},
		{
			name:   "conditional assignment",
			source: "Option Explicit\nPublic Sub Work()\nDim target As Excel.Worksheet\nDim source As Excel.Worksheet\n#If VBA7 Then\ntarget = source\n#End If\nEnd Sub\n",
		},
		{
			name:   "conditional declaration",
			source: "Option Explicit\nPublic Sub Work()\nDim source As Excel.Worksheet\n#If VBA7 Then\nDim target As Excel.Worksheet\n#End If\ntarget = source\nEnd Sub\n",
		},
		{
			name:   "continued assignment",
			source: "Option Explicit\nPublic Sub Work()\nDim target As Excel.Worksheet\nDim source As Excel.Worksheet\ntarget = _\n    source\nEnd Sub\n",
		},
		{
			name:   "member target",
			source: "Option Explicit\nPublic Sub Work()\nDim target As Excel.Worksheet\nDim source As Excel.Worksheet\ntarget.Name = source.Name\nEnd Sub\n",
		},
		{
			name:   "array target",
			source: "Option Explicit\nPublic Sub Work()\nDim targets(1) As Excel.Worksheet\nDim source As Excel.Worksheet\ntargets(0) = source\nEnd Sub\n",
		},
		{
			name:   "function target",
			source: "Option Explicit\nPublic Function GetTarget() As Excel.Worksheet\nEnd Function\nPublic Sub Work()\nDim source As Excel.Worksheet\nGetTarget() = source\nEnd Sub\n",
		},
		{
			name:   "function return slot target",
			source: "Option Explicit\nPublic Function GetTarget(ByVal source As Excel.Worksheet) As Excel.Worksheet\nGetTarget = source\nEnd Function\n",
		},
		{
			name:   "member expression RHS",
			source: "Option Explicit\nPublic Sub Work()\nDim target As Excel.Worksheet\nDim source As Excel.Worksheet\ntarget = source.Parent\nEnd Sub\n",
		},
		{
			name:   "array expression RHS",
			source: "Option Explicit\nPublic Sub Work()\nDim target As Excel.Worksheet\nDim sources(1) As Excel.Worksheet\ntarget = sources(0)\nEnd Sub\n",
		},
		{
			name:   "function expression RHS",
			source: "Option Explicit\nPublic Function GetSource() As Excel.Worksheet\nEnd Function\nPublic Sub Work()\nDim target As Excel.Worksheet\ntarget = GetSource()\nEnd Sub\n",
		},
		{
			name:   "compound string expression RHS",
			source: "Option Explicit\nPublic Sub Work()\nDim count As Long\nDim text As String\nSet count = \"a\" & text & \"b\"\nEnd Sub\n",
		},
		{
			name:   "Variant target",
			source: "Option Explicit\nPublic Sub Work()\nDim target As Variant\nDim source As Excel.Worksheet\ntarget = source\nEnd Sub\n",
		},
		{
			name:   "Object target",
			source: "Option Explicit\nPublic Sub Work()\nDim target As Object\nDim source As Excel.Worksheet\ntarget = source\nEnd Sub\n",
		},
		{
			name:   "implicitly Variant target",
			source: "Option Explicit\nPublic Sub Work()\nDim target\nDim source As Excel.Worksheet\ntarget = source\nEnd Sub\n",
		},
		{
			name:   "unresolved target type",
			source: "Option Explicit\nPublic Sub Work()\nDim target As Test.Missing\nDim source As Test.Missing\ntarget = source\nEnd Sub\n",
		},
		{
			name:   "unresolved RHS identifier",
			source: "Option Explicit\nPublic Sub Work()\nDim target As Excel.Worksheet\ntarget = missingSource\nEnd Sub\n",
		},
		{
			name:   "ambiguous local binding",
			source: "Option Explicit\nPublic Sub Work()\nDim target As Excel.Worksheet\nDim target As Excel.Worksheet\nDim source As Excel.Worksheet\ntarget = source\nEnd Sub\n",
		},
		{
			name:        "known incompatible object types",
			source:      "Option Explicit\nPublic Sub Work()\nDim target As Test.Control\nDim source As Test.Other\ntarget = source\nEnd Sub\n",
			customTypes: true,
			relation:    "known-incompatible",
		},
		{
			name:        "unknown object assignability",
			source:      "Option Explicit\nPublic Sub Work()\nDim target As Test.Control\nDim source As Test.Unrelated\ntarget = source\nEnd Sub\n",
			customTypes: true,
			relation:    "unknown",
		},
		{
			name:   "unrelated line selection",
			source: "Option Explicit\nPublic Sub Work()\nDim target As Excel.Worksheet\nDim source As Excel.Worksheet\ntarget = source\nEnd Sub\nPublic Sub Other()\n    Debug.Print \"unrelated\"\nEnd Sub\n",
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			analyzer := newTestAnalyzer(t)
			if tc.customTypes {
				addSetAssignmentTestTypes(t, &analyzer)
				assignable, known := analyzer.DB.IsAssignable("Test.Control", "Test.Other")
				if tc.relation == "known-incompatible" && (!known || assignable) {
					t.Fatalf("Test.Control <- Test.Other = %v, %v; want known incompatibility", assignable, known)
				}
				if tc.relation == "unknown" {
					assignable, known = analyzer.DB.IsAssignable("Test.Control", "Test.Unrelated")
					if known || assignable {
						t.Fatalf("Test.Control <- Test.Unrelated = %v, %v; want unknown assignability", assignable, known)
					}
				}
			}
			doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: tc.source}
			selection := Range{End: positionForByteOffset(doc.Source, len(doc.Source))}
			if tc.name == "unrelated line selection" {
				line := strings.Index(doc.Source, "Debug.Print")
				selection = Range{Start: positionForByteOffset(doc.Source, line), End: positionForByteOffset(doc.Source, line+len("Debug.Print \"unrelated\""))}
			}
			assertNoSetAssignmentFix(t, analyzer, doc, selection)
		})
	}
}

func TestSetAssignmentQuickFixMatchesVBEFixtures(t *testing.T) {
	for _, tc := range []struct{ before, after, code string }{
		{"quickfix-set-object-before", "quickfix-set-object-after", "VB036"},
		{"set-scalar-target", "scalar-assignment", "VB037"},
	} {
		t.Run(tc.code, func(t *testing.T) {
			root := filepath.Join("..", "..", "..", "testdata", "vbe-oracle", "cases")
			before, err := os.ReadFile(filepath.Join(root, tc.before, "Main.bas"))
			if err != nil {
				t.Fatal(err)
			}
			after, err := os.ReadFile(filepath.Join(root, tc.after, "Main.bas"))
			if err != nil {
				t.Fatal(err)
			}
			analyzer := newTestAnalyzer(t)
			doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: string(before)}
			actions, err := analyzer.CodeActionsContext(t.Context(), doc, Range{End: positionForByteOffset(doc.Source, len(doc.Source))}, []string{"quickfix"})
			if err != nil {
				t.Fatal(err)
			}
			fixes := setAssignmentActionsByCode(actions, tc.code)
			if len(fixes) != 1 {
				t.Fatalf("actions = %+v", actions)
				return
			}
			want := string(after)
			if tc.code == "VB037" {
				// Existing controls use different procedure names; their bodies are
				// otherwise the exact rejected/accepted assignment pair.
				want = strings.ReplaceAll(want, "ProbeScalarAssignment", "ProbeSetScalarTarget")
			}
			if got := applyQuickFixForTest(t, doc, fixes[0]); got != want {
				t.Fatalf("corrected = %q, want VBE control %q", got, want)
			}
		})
	}
}

func TestSetAssignmentQuickFixRequiresPresentCompleteTypeDB(t *testing.T) {
	for _, state := range []string{"missing", "incomplete"} {
		t.Run(state, func(t *testing.T) {
			analyzer := newTestAnalyzer(t)
			if state == "missing" {
				analyzer.DB = nil
			} else {
				analyzer.TypeDBResolutionIncomplete = true
			}
			source := "Option Explicit\nPublic Sub Work()\nDim target As Excel.Worksheet\nDim source As Excel.Worksheet\ntarget = source\nEnd Sub\n"
			doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: source}
			assertNoSetAssignmentFix(t, analyzer, doc, Range{End: positionForByteOffset(doc.Source, len(doc.Source))})
		})
	}
}

func TestSetAssignmentQuickFixFiltersSuppressionAcrossSelection(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := "Option Explicit\nSub Work()\nDim target As Excel.Worksheet\nDim source As Excel.Worksheet\nDim count As Long\ntarget = source ' xlflow:disable-line VB036\ntarget = source\nSet count = 1 ' xlflow:disable-line VB037\nEnd Sub\n"
	doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: source}
	actions, err := analyzer.CodeActionsContext(t.Context(), doc, Range{End: positionForByteOffset(source, len(source))}, []string{"quickfix"})
	if err != nil {
		t.Fatal(err)
	}
	if len(actions) != 2 || actions[0].Diagnostics[0].Code != "VB036" || actions[0].Diagnostics[0].Range.Start.Line != 6 || actions[1].Diagnostics[0].Code != "VB037" {
		t.Fatalf("visible actions = %+v", actions)
	}
}

func TestSetAssignmentQuickFixHonorsDiagnosticVisibilityAndSuppression(t *testing.T) {
	tests := []struct {
		name, source, code string
		wantAction         bool
	}{
		{
			name:   "suppressed VB036",
			source: "Option Explicit\nPublic Sub Work()\nDim target As Excel.Worksheet\nDim source As Excel.Worksheet\ntarget = source ' xlflow:disable-line VB036\nEnd Sub\n",
			code:   "VB036",
		},
		{
			name:   "VB037 remains visible under suppression",
			source: "Option Explicit\nPublic Sub Work()\nDim count As Long\nSet count = 1 ' xlflow:disable-line VB037\nEnd Sub\n",
			code:   "VB037", wantAction: true,
		},
		{
			name:   "Let tab has no existing VB036 diagnostic",
			source: "Option Explicit\nPublic Sub Work()\nDim target As Excel.Worksheet\nDim source As Excel.Worksheet\nLet\ttarget = source\nEnd Sub\n",
			code:   "VB036",
		},
		{
			name:   "Set tab has no existing VB037 diagnostic",
			source: "Option Explicit\nPublic Sub Work()\nDim count As Long\nSet\tcount = 1\nEnd Sub\n",
			code:   "VB037",
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			analyzer := newTestAnalyzer(t)
			doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: tc.source}
			diagnostics := diagnosticsByCode(analyzer.Diagnostics(doc), tc.code)
			if tc.wantAction && len(diagnostics) != 1 {
				t.Fatalf("visible %s diagnostics = %+v, want one", tc.code, diagnostics)
			}
			if !tc.wantAction && len(diagnostics) != 0 {
				t.Fatalf("%s diagnostics = %+v, want none", tc.code, diagnostics)
			}
			actions, err := analyzer.CodeActionsContext(t.Context(), doc, Range{End: positionForByteOffset(doc.Source, len(doc.Source))}, []string{"quickfix"})
			if err != nil {
				t.Fatal(err)
			}
			setActions := setAssignmentActionsByCode(actions, tc.code)
			if tc.wantAction && len(setActions) != 1 {
				t.Fatalf("%s actions = %+v, want one", tc.code, setActions)
			}
			if !tc.wantAction && len(setActions) != 0 {
				t.Fatalf("%s actions = %+v, want none", tc.code, setActions)
			}
			if tc.wantAction && !reflect.DeepEqual(setActions[0].Diagnostics, diagnostics) {
				t.Fatalf("VB037 associated diagnostics = %+v, published = %+v", setActions[0].Diagnostics, diagnostics)
			}
		})
	}
}

func TestSetAssignmentQuickFixSelectionStaysLocalAcrossUnrelatedProcedures(t *testing.T) {
	var source strings.Builder
	source.WriteString("Option Explicit\nPublic Sub TargetProcedure()\nDim target As Excel.Worksheet\nDim source As Excel.Worksheet\ntarget = source\nEnd Sub\n")
	for i := range 80 {
		source.WriteString("Public Sub Unrelated")
		source.WriteString(strconv.Itoa(i))
		source.WriteString("()\nDebug.Print ")
		source.WriteString(string(rune('0' + i%10)))
		source.WriteString("\nEnd Sub\n")
	}
	analyzer := newTestAnalyzer(t)
	doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: source.String()}
	snapshot := NewAnalysisSnapshot(doc)
	defer snapshot.Retire()
	doc = snapshot.Document()
	if _, err := analyzer.CodeActionsContext(t.Context(), doc, Range{}, []string{"quickfix"}); err != nil {
		t.Fatal(err)
	}
	if snapshot.index != nil {
		t.Fatal("non-assignment selection built the type index")
	}
	targetOffset := strings.Index(doc.Source, "target = source")
	selection := Range{Start: positionForByteOffset(doc.Source, targetOffset), End: positionForByteOffset(doc.Source, targetOffset+len("target = source"))}
	actions, err := analyzer.CodeActionsContext(t.Context(), doc, selection, []string{"quickfix"})
	if err != nil {
		t.Fatal(err)
	}
	if got := setAssignmentActionsByCode(actions, "VB036"); len(got) != 1 {
		t.Fatalf("target selection actions = %+v, want one VB036", got)
	}
	if snapshot.index == nil {
		t.Fatal("positive assignment selection did not build the type index")
	}
	if stats := snapshot.ProcedureArtifactStats(); stats.IRBuild != 0 || stats.CFGBuild != 0 {
		t.Fatalf("selected correction built procedure artifacts: %+v", stats)
	}

	unrelatedBase := targetOffset + len("target = source")
	unrelatedIndex := strings.Index(doc.Source[unrelatedBase:], "Debug.Print")
	if unrelatedIndex < 0 {
		t.Fatal("unrelated procedure has no selection target")
	}
	unrelatedOffset := unrelatedBase + unrelatedIndex
	unrelatedSelection := Range{Start: positionForByteOffset(doc.Source, unrelatedOffset), End: positionForByteOffset(doc.Source, unrelatedOffset+len("Debug.Print 0"))}
	actions, err = analyzer.CodeActionsContext(t.Context(), doc, unrelatedSelection, []string{"quickfix"})
	if err != nil {
		t.Fatal(err)
	}
	if got := setAssignmentActionsByCode(actions, "VB036"); len(got) != 0 {
		t.Fatalf("unrelated selection actions = %+v, want none", got)
	}
}

func assertSetAssignmentQuickFix(t *testing.T, analyzer Analyzer, source, want, code, title, editAnchor string, editOffset, editLength int, newText string) {
	t.Helper()
	doc := Document{Path: filepath.Join(analyzer.RootDir, "Main.bas"), Source: source, Version: 1}
	snapshot := NewAnalysisSnapshot(doc)
	defer snapshot.Retire()
	doc = snapshot.Document()
	selection := Range{End: positionForDocumentByteOffset(doc, len(source))}
	actions, err := analyzer.CodeActionsContext(t.Context(), doc, selection, []string{"quickfix"})
	if err != nil {
		t.Fatal(err)
	}
	setActions := setAssignmentActionsByCode(actions, code)
	if len(setActions) != 1 {
		t.Fatalf("%s actions = %+v, want one", code, setActions)
		return
	}
	if stats := snapshot.ProcedureArtifactStats(); stats.IRBuild != 0 || stats.CFGBuild != 0 {
		t.Fatalf("quick-fix discovery built procedure artifacts before Diagnostics: %+v", stats)
	}
	action := setActions[0]
	if action.Title != title || action.Kind != "quickfix" {
		t.Fatalf("action = %+v, want title %q and quickfix kind", action, title)
	}
	start := strings.Index(source, editAnchor)
	if start < 0 {
		t.Fatalf("edit anchor %q missing from source", editAnchor)
	}
	start += editOffset
	wantEdit := CodeActionEdit{
		Range: Range{
			Start: positionForDocumentByteOffset(doc, start),
			End:   positionForDocumentByteOffset(doc, start+editLength),
		},
		NewText: newText,
	}
	if !reflect.DeepEqual(action.Edits, []CodeActionEdit{wantEdit}) {
		t.Fatalf("edits = %+v, want %+v", action.Edits, []CodeActionEdit{wantEdit})
	}

	published := diagnosticsByCode(analyzer.Diagnostics(doc), code)
	if len(published) != 1 {
		t.Fatalf("published %s diagnostics = %+v, want one", code, published)
		return
	}
	if !reflect.DeepEqual(action.Diagnostics, []Diagnostic{published[0]}) {
		t.Fatalf("associated diagnostics = %+v, want full published diagnostic %+v", action.Diagnostics, published[0])
	}

	corrected := applyQuickFixForTest(t, doc, action)
	if corrected != want {
		t.Fatalf("corrected source = %q, want %q", corrected, want)
	}
	correctedDoc := Document{Path: doc.Path, Source: corrected}
	if diagnostics := diagnosticsByCode(analyzer.Diagnostics(correctedDoc), code); len(diagnostics) != 0 {
		t.Fatalf("corrected source still has %s diagnostics: %+v", code, diagnostics)
	}
	actions, err = analyzer.CodeActionsContext(t.Context(), correctedDoc, Range{End: positionForByteOffset(corrected, len(corrected))}, []string{"quickfix"})
	if err != nil {
		t.Fatal(err)
	}
	if again := setAssignmentActionsByCode(actions, code); len(again) != 0 {
		t.Fatalf("second %s quick fix = %+v, want none", code, again)
	}
}

func assertNoSetAssignmentFix(t *testing.T, analyzer Analyzer, doc Document, selection Range) {
	t.Helper()
	actions, err := analyzer.CodeActionsContext(t.Context(), doc, selection, []string{"quickfix"})
	if err != nil {
		t.Fatal(err)
	}
	for _, code := range []string{"VB036", "VB037"} {
		if got := setAssignmentActionsByCode(actions, code); len(got) != 0 {
			t.Errorf("unexpected %s quick fixes: %+v", code, got)
		}
	}
}

func setAssignmentActionsByCode(actions []CodeAction, code string) []CodeAction {
	var out []CodeAction
	for _, action := range actions {
		for _, diagnostic := range action.Diagnostics {
			if diagnostic.Code == code {
				out = append(out, action)
				break
			}
		}
	}
	return out
}

func addSetAssignmentTestTypes(t *testing.T, analyzer *Analyzer) {
	t.Helper()
	if err := analyzer.DB.MergeJSON([]byte(`{
  "types": [
    { "name": "Test.Control", "library": "Test", "kind": "interface" },
    { "name": "Test.TextBox", "library": "Test", "kind": "interface", "assignable_to": ["Test.Control"] },
    { "name": "Test.Marker", "library": "Test", "kind": "interface" },
    { "name": "Test.Other", "library": "Test", "kind": "interface", "assignable_to": ["Test.Marker"] }
  ]
}`)); err != nil {
		t.Fatal(err)
	}
}
