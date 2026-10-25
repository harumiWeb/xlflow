package intel

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	staticrules "github.com/harumiWeb/xlflow/internal/staticanalysis/rules"
	"github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func byValFixTestAnalyzer(t *testing.T, source string) (Analyzer, Document, Diagnostic) {
	t.Helper()
	a := newTestAnalyzer(t)
	a.DB = nil
	a.Config.Lint.RequireOptionExplicit = false
	a.Config.Lint.DetectUnusedLocalVariables = false
	a.Config.Analyze.DetectByRefParametersCanBeByVal = true
	doc := Document{Path: "Main.bas", URI: "file:///Main.bas", ModuleKind: "standard", Source: source, Version: 23}
	start := strings.Index(strings.ToLower(source), "byref")
	if start < 0 {
		start = 0
	}
	diagnostic := Diagnostic{Code: "VBA275", Severity: "information", Source: "xlflow", Message: "current analyzer diagnostic", Range: Range{Start: positionForByteOffset(source, start), End: positionForByteOffset(source, start+5)}}
	a.QuickFixDiagnosticsFunc = func(current Document) []Diagnostic {
		if current.Source != doc.Source || current.Version != doc.Version || current.Path != doc.Path {
			return nil
		}
		return []Diagnostic{diagnostic}
	}
	return a, doc, diagnostic
}

func TestByValQuickFixPreservesSourceAndDiagnostic(t *testing.T) {
	for _, typ := range []string{"Boolean", "Byte", "Integer", "Long", "LongLong", "LongPtr", "Single", "Double", "Currency", "Date", "String"} {
		for _, eol := range []string{"\n", "\r\n"} {
			for _, body := range []string{"    Show = value\n", "    Show = (value + 1)\n", "    ' 日本語 😀\n    Show = value\n"} {
				t.Run(typ+eol+body, func(t *testing.T) {
					source := strings.ReplaceAll("Option Explicit\nPrivate Function Show(\tByRef\tvalue As "+typ+") As "+typ+" ' keep 日本語😀\n"+body+"End Function\n", "\n", eol)
					a, doc, diagnostic := byValFixTestAnalyzer(t, source)
					snapshot := NewAnalysisSnapshot(doc)
					t.Cleanup(snapshot.RetireAndWait)
					doc = snapshot.Document()
					actions, err := a.CodeActionsContext(t.Context(), doc, diagnostic.Range, []string{"quickfix"})
					if err != nil || len(actions) != 1 {
						parsed, _ := snapshot.ParsedDocument()
						_ = parsed.Read(func(view ast.ParsedView) error { t.Log(view.Root.ToSexp()); return nil })
						t.Fatalf("actions = %+v, err = %v", actions, err)
					}
					if !reflect.DeepEqual(actions[0].Diagnostics, []Diagnostic{diagnostic}) {
						t.Fatalf("associated diagnostic = %+v", actions[0].Diagnostics)
					}
					if got, want := applyQuickFixForTest(t, doc, actions[0]), strings.Replace(source, "ByRef", "ByVal", 1); got != want {
						t.Fatalf("corrected source = %q, want %q", got, want)
					}
					if stats := snapshot.ProcedureArtifactStats(); stats.IRBuild != 0 || stats.CFGBuild != 0 {
						t.Fatalf("discovery constructed artifacts: %+v", stats)
					}
				})
			}
		}
	}
}

func TestByValQuickFixFunctions(t *testing.T) {
	for _, assignment := range []string{
		"Adjust = value + 1", "Adjust = value - 1", "Let Adjust = (value * 2)",
		"Adjust = value / 2", "Adjust = value \\ 2", "Adjust = value Mod 2", "Adjust = value mOd 2",
		"Adjust = value ^ 2", "Adjust = -(value + 1)", "Adjust = +value", "Adjust = Not value",
		"Adjust = \"ByRef 日本語😀\"",
	} {
		source := "Option Explicit\nPrivate Function Adjust(ByRef value As Long) As String\n    " + assignment + "\nEnd Function\n"
		a, doc, diagnostic := byValFixTestAnalyzer(t, source)
		snapshot := NewAnalysisSnapshot(doc)
		t.Cleanup(snapshot.RetireAndWait)
		doc = snapshot.Document()
		actions, err := a.CodeActionsContext(t.Context(), doc, diagnostic.Range, []string{"quickfix"})
		if err != nil || len(actions) != 1 {
			parsed, _ := snapshot.ParsedDocument()
			_ = parsed.Read(func(view ast.ParsedView) error { t.Log(view.Root.ToSexp()); return nil })
			t.Fatalf("%s: actions = %+v, err = %v", assignment, actions, err)
		}
	}
}

func TestByValQuickFixWithholdsUnsafeFormsEvenWithDiagnostic(t *testing.T) {
	for _, tc := range []struct{ name, before, header, body, end, kind string }{
		{name: "public", header: "Public Function Show(ByRef value As Long)"},
		{name: "implicit visibility", header: "Function Show(ByRef value As Long)"},
		{name: "friend", header: "Friend Function Show(ByRef value As Long)"},
		{name: "static", header: "Private Static Function Show(ByRef value As Long)"},
		{name: "sub", header: "Private Sub Show(ByRef value As Long)", body: "Debug.Print value\n", end: "End Sub\n"},
		{name: "optional", header: "Private Function Show(Optional ByRef value As Long = 0)"},
		{name: "implicit passing", header: "Private Function Show(value As Long)"},
		{name: "array", header: "Private Function Show(ByRef value() As Long)"},
		{name: "object", header: "Private Function Show(ByRef value As Object)"},
		{name: "variant", header: "Private Function Show(ByRef value As Variant)"},
		{name: "UDT", header: "Private Function Show(ByRef value As Pair)"},
		{name: "multiple aliases", header: "Private Function Show(ByRef value As Long, ByRef other As Long)", body: "    other = 4\n    Show = value\n"},
		{name: "module write alias", before: "Private shared As Long\n", body: "    shared = 4\n    Show = value\n"},
		{name: "global read", before: "Private shared As Long\n", body: "    Debug.Print shared, value\n"},
		{name: "parameter write", body: "    value = 4\n"},
		{name: "call", body: "    Mutate value\n"},
		{name: "bare call", body: "    DoEvents\n    Show = value\n"},
		{name: "call expression", body: "    Debug.Print GetValue(value)\n"},
		{name: "member", body: "    Show = value.Count\n"},
		{name: "local binding", body: "    Dim local As Long\n    local = value\n"},
		{name: "flow", body: "    If value > 0 Then Show = value\n"},
		{name: "with", body: "    With value\n    End With\n"},
		{name: "continued header", header: "Private Function Show( _\n    ByRef value As Long)"},
		{name: "continued body", body: "    Show = value + _\n    1\n"},
		{name: "colon", body: "    Show = value: Show = value\n"},
		{name: "conditional", before: "#If VBA7 Then\n", end: "End Function\n#End If\n"},
		{name: "conditional body", body: "#If VBA7 Then\nShow = value\n#End If\n"},
		{name: "attribute", body: "Attribute Show.VB_Description = \"test\"\nShow = value\n"},
		{name: "duplicate", before: "Private Function Show() As Long\nEnd Function\n"},
		{name: "class", kind: "class"},
		{name: "form", kind: "form"},
		{name: "document", kind: "document"},
		{name: "unknown module", kind: "unknown"},
		{name: "malformed", body: "    Debug.Print (value\n"},
		{name: "empty body", body: "\n"},
		{name: "comment only", body: "' only documentation\n"},
		{name: "return object", header: "Private Function Show(ByRef value As Long) As Object"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			header, body, end := tc.header, tc.body, tc.end
			if header == "" {
				header = "Private Function Show(ByRef value As Long)"
			}
			if strings.Contains(header, "Function ") && strings.HasSuffix(header, ")") {
				header += " As Long"
			}
			if body == "" {
				body = "    Show = value\n"
			}
			if end == "" {
				end = "End Function\n"
			}
			a, doc, diagnostic := byValFixTestAnalyzer(t, "Option Explicit\n"+tc.before+header+"\n"+body+end)
			if tc.kind != "" {
				doc.ModuleKind = tc.kind
			}
			actions, err := a.CodeActionsContext(t.Context(), doc, diagnostic.Range, []string{"quickfix"})
			if err != nil || len(actions) != 0 {
				t.Fatalf("actions = %+v, err = %v", actions, err)
			}
		})
	}
}

func TestByValQuickFixWithholdsNonArithmeticBinaryExpressions(t *testing.T) {
	for _, expression := range []string{
		"value And True", "value Or False", "value Xor True", "value Eqv True", "value Imp True",
		"value & \"suffix\"", "value + (value And 1)", "-(value Or 1)",
	} {
		t.Run(expression, func(t *testing.T) {
			source := "Private Function Show(ByRef value As Long) As Long\n    Show = " + expression + "\nEnd Function\n"
			a, doc, diagnostic := byValFixTestAnalyzer(t, source)
			actions, err := a.CodeActionsContext(t.Context(), doc, diagnostic.Range, []string{"quickfix"})
			if err != nil || len(actions) != 0 {
				t.Fatalf("non-arithmetic expression offered a correction: actions=%+v err=%v", actions, err)
			}
		})
	}
}

func TestByValQuickFixRequiresCurrentEvidenceAndSelection(t *testing.T) {
	const source = "Private Function Show(ByRef value As Long) As Long\nShow = value\nEnd Function\n"
	for _, mode := range []string{"missing", "empty", "disabled", "old version", "wrong range", "wrong code", "unrelated", "canceled", "refactor only", "fix all"} {
		t.Run(mode, func(t *testing.T) {
			a, doc, diagnostic := byValFixTestAnalyzer(t, source)
			ctx := t.Context()
			only := []string{"quickfix"}
			selection := diagnostic.Range
			switch mode {
			case "missing":
				a.QuickFixDiagnosticsFunc = nil
			case "empty":
				a.QuickFixDiagnosticsFunc = func(Document) []Diagnostic { return nil }
			case "disabled":
				a.Config.Analyze.DetectByRefParametersCanBeByVal = false
			case "old version":
				doc.Version++
			case "wrong range":
				diagnostic.Range.End.Character++
				a.QuickFixDiagnosticsFunc = func(Document) []Diagnostic { return []Diagnostic{diagnostic} }
			case "wrong code":
				diagnostic.Code = "VBA274"
				a.QuickFixDiagnosticsFunc = func(Document) []Diagnostic { return []Diagnostic{diagnostic} }
			case "unrelated":
				selection = Range{}
			case "canceled":
				var cancel context.CancelFunc
				ctx, cancel = context.WithCancel(ctx)
				cancel()
			case "refactor only":
				only = []string{"refactor.rewrite"}
			case "fix all":
				only = []string{"source.fixAll.xlflow.vba275"}
			}
			actions, err := a.CodeActionsContext(ctx, doc, selection, only)
			if mode == "canceled" && err != context.Canceled {
				t.Fatalf("err = %v", err)
			}
			for _, action := range actions {
				if len(action.Diagnostics) != 0 && action.Diagnostics[0].Code == "VBA275" {
					t.Fatalf("actions = %+v", actions)
				}
			}
		})
	}
}

func TestByValQuickFixDiscoveryDoesNotBuildUnrelatedArtifacts(t *testing.T) {
	source := "Private Function Show(ByRef value As Long) As Long\nShow = value\nEnd Function\n"
	for i := range 128 {
		source += fmt.Sprintf("Private Sub Unrelated%d()\nEnd Sub\n", i)
	}
	a, doc, diagnostic := byValFixTestAnalyzer(t, source)
	snapshot := NewAnalysisSnapshot(doc)
	t.Cleanup(snapshot.RetireAndWait)
	doc = snapshot.Document()
	a.WorkspaceSymbolsFunc = func([]Document, string) ([]Symbol, error) { t.Fatal("workspace symbols loaded"); return nil, nil }
	a.RealtimeFindingsFunc = func(context.Context, string, config.Config, *ast.ParsedDocument, procedureir.DocumentIR, cfg.Document) ([]RealtimeFinding, error) {
		t.Fatal("analysis invoked")
		return nil, nil
	}
	actions, err := a.CodeActionsContext(t.Context(), doc, diagnostic.Range, []string{"quickfix"})
	if err != nil || len(actions) != 1 {
		t.Fatalf("actions = %+v, err = %v", actions, err)
	}
	if stats := snapshot.ProcedureArtifactStats(); stats.IRBuild != 0 || stats.CFGBuild != 0 {
		t.Fatalf("artifact builds: %+v", stats)
	}
	if builds, hits := snapshot.InteractiveIndexStats(); builds != 0 || hits != 0 {
		t.Fatal("interactive index constructed")
	}
}

func TestByValQuickFixMatchesVBEFixtures(t *testing.T) {
	root := filepath.Join("..", "..", "..", "testdata", "vbe-oracle", "cases")
	before, err := os.ReadFile(filepath.Join(root, "quickfix-byval-before", "Main.bas"))
	if err != nil {
		t.Fatal(err)
	}
	after, err := os.ReadFile(filepath.Join(root, "quickfix-byval-after", "Main.bas"))
	if err != nil {
		t.Fatal(err)
	}
	a, doc, diagnostic := byValFixTestAnalyzer(t, string(before))
	var published []Diagnostic
	for offset := 0; ; {
		i := strings.Index(doc.Source[offset:], "ByRef")
		if i < 0 {
			break
		}
		start := offset + i
		current := diagnostic
		current.Range = Range{Start: positionForByteOffset(doc.Source, start), End: positionForByteOffset(doc.Source, start+5)}
		published = append(published, current)
		offset = start + 5
	}
	a.QuickFixDiagnosticsFunc = func(Document) []Diagnostic { return published }
	selection := Range{End: positionForByteOffset(doc.Source, len(doc.Source))}
	actions, err := a.CodeActionsContext(t.Context(), doc, selection, []string{"quickfix"})
	if err != nil || len(actions) != 4 {
		t.Fatalf("fixture actions = %+v, err = %v", actions, err)
	}
	combined := CodeAction{}
	for _, action := range actions {
		combined.Edits = append(combined.Edits, action.Edits...)
	}
	if got := applyQuickFixForTest(t, doc, combined); got != string(after) {
		t.Fatalf("fixture edit differs: %q", got)
	}
	doc.Source = string(after)
	actions, err = a.CodeActionsContext(t.Context(), doc, selection, []string{"quickfix"})
	if err != nil || len(actions) != 0 {
		t.Fatalf("second application = %+v, err = %v", actions, err)
	}
}

func TestByValQuickFixRemainsOptInAndIndividualOnly(t *testing.T) {
	rule, ok := staticrules.Lookup("VBA275")
	if !ok || !rule.FixAvailable || rule.DefaultEnabled || rule.DefaultSeverity != staticrules.SeverityInformation || !rule.InlineSuppressible || rule.CompileEquivalent {
		t.Fatalf("VBA275 metadata = %+v", rule)
	}
	for _, code := range documentFixAllRules {
		if code == "VBA275" {
			t.Fatal("VBA275 unexpectedly eligible for Fix All")
		}
	}
	if config.Default().Analyze.DetectByRefParametersCanBeByVal {
		t.Fatal("VBA275 enabled by default")
	}
}

func TestByValQuickFixWithoutEvidenceDoesNotParse(t *testing.T) {
	a, doc, diagnostic := byValFixTestAnalyzer(t, "Private Function Echo(ByRef value As Long) As Long\nEcho = value\nEnd Function\n")
	a.QuickFixDiagnosticsFunc = func(Document) []Diagnostic { return nil }
	snapshot := NewAnalysisSnapshot(doc)
	t.Cleanup(snapshot.RetireAndWait)
	doc = snapshot.Document()
	actions, err := a.CodeActionsContext(t.Context(), doc, diagnostic.Range, []string{"quickfix"})
	if err != nil || len(actions) != 0 || snapshot.ParseCount() != 0 {
		t.Fatalf("missing evidence parsed or offered an action: actions=%+v err=%v parses=%d", actions, err, snapshot.ParseCount())
	}
}
